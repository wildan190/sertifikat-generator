import 'dotenv/config'
import { ZipArchive } from 'archiver'
import cors from 'cors'
import express from 'express'
import fs from 'fs'
import mongoose, { Schema } from 'mongoose'
import multer from 'multer'
import os from 'os'
import path from 'path'
import PDFDocument from 'pdfkit'
import { randomUUID } from 'crypto'
import * as XLSX from 'xlsx'

const app = express()
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 8 * 1024 * 1024 } })
app.use(cors())
app.use(express.json({ limit: '20mb' }))

type TemplatePayload = {
  name: string; title: string; subtitle?: string; type: string; primaryField: string; body: string
  signatory: string; organization: string; accent: string; logo?: string
  logo2?: string
  enableLogo2?: boolean
  enableSig2?: boolean
  signatureImage?: string; signatureImage2?: string
  signatory2?: string; signatoryTitle1?: string; signatoryTitle2?: string
  certNumberPrefix?: string
  certificateDate?: string
}

const templateSchema = new Schema<TemplatePayload>({
  name: { type: String, required: true }, title: { type: String, required: true },
  subtitle: { type: String, default: '' },
  type: { type: String, required: true }, primaryField: { type: String, required: true },
  body: { type: String, required: true }, signatory: { type: String, required: true },
  signatory2: { type: String, default: '' },
  signatoryTitle1: { type: String, default: '' },
  signatoryTitle2: { type: String, default: '' },
  organization: { type: String, required: true }, accent: { type: String, required: true },
  logo: { type: String, default: '' },
  logo2: { type: String, default: '' },
  enableLogo2: { type: Boolean, default: false },
  enableSig2: { type: Boolean, default: false },
  signatureImage: { type: String, default: '' },
  signatureImage2: { type: String, default: '' },
  certNumberPrefix: { type: String, default: 'CERT' },
  certificateDate: { type: String, default: '' },
}, { timestamps: true })
const CertificateTemplate = mongoose.model<TemplatePayload>('CertificateTemplate', templateSchema)

let databaseState: 'connecting' | 'connected' | 'offline' = 'connecting'
let cachedPromise: Promise<typeof mongoose> | null = null

async function connectToDatabase() {
  const uri = process.env.MONGODB_URI
  if (!uri) {
    databaseState = 'offline'
    return null
  }
  if (mongoose.connection.readyState >= 1) {
    databaseState = 'connected'
    return mongoose
  }
  if (!cachedPromise) {
    cachedPromise = mongoose.connect(uri, { dbName: 'certificate_generator', serverSelectionTimeoutMS: 5000 })
      .then((m) => {
        databaseState = 'connected'
        return m
      })
      .catch((err) => {
        databaseState = 'offline'
        cachedPromise = null
        throw err
      })
  }
  return cachedPromise
}

connectToDatabase().catch(() => {})

const value = (row: Record<string, unknown>, key: string) => String(row[key] ?? '')
const interpolate = (source: string, row: Record<string, unknown>, template?: TemplatePayload) =>
  source.replace(/{{\s*([^}]+)\s*}}/g, (_, key: string) => {
    const trimmed = key.trim()
    const val = value(row, trimmed)
    if (val) return val
    if ((trimmed.toLowerCase() === 'tanggal' || trimmed.toLowerCase() === 'date') && template?.certificateDate) {
      return template.certificateDate
    }
    return `{{${trimmed}}}`
  })

app.get('/api/health', async (_req, res) => {
  try { await connectToDatabase() } catch { /* ignore */ }
  res.json({ database: databaseState })
})

app.get('/api/templates', async (_req, res) => {
  try { await connectToDatabase() } catch { /* ignore */ }
  if (databaseState !== 'connected') return res.status(503).json({ message: 'MongoDB belum tersambung.' })
  res.json(await CertificateTemplate.find().sort({ updatedAt: -1 }).lean())
})

app.post('/api/templates', async (req, res) => {
  try { await connectToDatabase() } catch { /* ignore */ }
  if (databaseState !== 'connected') return res.status(503).json({ message: 'MongoDB belum tersambung.' })
  const payload = req.body as TemplatePayload
  if (!payload.name || !payload.primaryField || !payload.body) return res.status(400).json({ message: 'Nama, kolom utama, dan isi sertifikat wajib diisi.' })
  res.status(201).json(await CertificateTemplate.create(payload))
})

app.delete('/api/templates/:id', async (req, res) => {
  try { await connectToDatabase() } catch { /* ignore */ }
  if (databaseState !== 'connected') return res.status(503).json({ message: 'MongoDB belum tersambung.' })
  try {
    const deleted = await CertificateTemplate.findByIdAndDelete(req.params.id)
    if (!deleted) return res.status(404).json({ message: 'Template tidak ditemukan.' })
    res.json({ message: 'Template berhasil dihapus.' })
  } catch {
    res.status(400).json({ message: 'ID template tidak valid.' })
  }
})

app.post('/api/import', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ message: 'File Excel belum dipilih.' })
  try {
    const book = XLSX.read(req.file.buffer, { type: 'buffer', raw: false })
    const sheet = book.Sheets[book.SheetNames[0] ?? '']
    if (!sheet) throw new Error('Sheet tidak ditemukan')
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' })
    res.json({ headers: rows.length ? Object.keys(rows[0] ?? {}) : [], rows })
  } catch {
    res.status(400).json({ message: 'File tidak dapat dibaca. Gunakan .xlsx atau .xls dengan header pada baris pertama.' })
  }
})

const getCertificateNumber = (template: TemplatePayload, recipient: Record<string, unknown>) => {
  for (const key of ['nomor', 'no_sertifikat', 'nomor_sertifikat', 'certificate_number', 'no']) {
    if (recipient[key]) return String(recipient[key]).trim()
  }
  const prefix = (template.certNumberPrefix || 'CERT').trim().toUpperCase()
  const rawId = `${value(recipient, template.primaryField)}_${template.title}_${template.signatory}`
  let hash = 0
  for (let i = 0; i < rawId.length; i++) {
    hash = ((hash << 5) - hash) + rawId.charCodeAt(i)
    hash |= 0
  }
  const positive = Math.abs(hash)
  const hexPart = positive.toString(16).toUpperCase().padStart(6, '0').slice(-6)
  const numPart = String(positive % 9000 + 1000)
  const year = new Date().getFullYear()
  return `${prefix}/${year}/${numPart}-${hexPart}`
}

function scallopSealPath(cx: number, cy: number, baseR: number, lobes = 16, bump = 2.5): string {
  const steps = lobes * 2
  const parts: string[] = []
  for (let i = 0; i <= steps; i++) {
    const angle = (i / steps) * Math.PI * 2 - Math.PI / 2
    const radius = i % 2 === 0 ? baseR + bump : baseR - bump * 0.35
    const x = cx + radius * Math.cos(angle)
    const y = cy + radius * Math.sin(angle)
    parts.push(i === 0 ? `M ${x.toFixed(2)} ${y.toFixed(2)}` : `L ${x.toFixed(2)} ${y.toFixed(2)}`)
  }
  return `${parts.join(' ')} Z`
}

function drawSealArcText(
  doc: PDFKit.PDFDocument,
  cx: number,
  cy: number,
  radius: number,
  text: string,
  color: string,
  position: 'top' | 'bottom'
) {
  const chars = text.split('')
  const arcSpan = Math.PI * 0.78
  const startAngle = position === 'top' ? -Math.PI / 2 - arcSpan / 2 : Math.PI / 2 - arcSpan / 2
  const step = chars.length > 1 ? arcSpan / (chars.length - 1) : 0

  doc.font('Helvetica-Bold').fontSize(5).fillColor(color)

  chars.forEach((char, index) => {
    const angle = startAngle + step * index
    const x = cx + radius * Math.cos(angle)
    const y = cy + radius * Math.sin(angle)
    const rotation = (angle * 180) / Math.PI + (position === 'top' ? 90 : -90)

    doc.save()
    doc.translate(x, y)
    doc.rotate(rotation)
    doc.text(char, -1.5, -2, { lineBreak: false, width: 4 })
    doc.restore()
  })
}

function drawOfficialSeal(doc: PDFKit.PDFDocument, cx: number, cy: number, accentColor: string, size: 'md' | 'lg' = 'md') {
  const outerR = size === 'lg' ? 30 : 27
  const bump = size === 'lg' ? 2.8 : 2.5
  const scallop = scallopSealPath(cx, cy, outerR, 16, bump)

  doc.save()

  doc.path(scallop).fillOpacity(0.14).fillColor('#d7c28a').fill()
  doc.path(scallop).fillOpacity(1).lineWidth(1.1).strokeColor('#d7c28a').stroke()

  doc.circle(cx, cy, outerR - 5).fillOpacity(0.95).fillColor('#ffffff').fill()
  doc.circle(cx, cy, outerR - 7).lineWidth(0.8).strokeColor('#d7c28a').stroke()
  doc.circle(cx, cy, outerR - 10).lineWidth(0.7).strokeColor(accentColor).stroke()

  const dotR = outerR - 2.5
  for (const [dx, dy] of [[0, -1], [1, 0], [0, 1], [-1, 0]] as const) {
    doc.circle(cx + dx * dotR * 0.82, cy + dy * dotR * 0.82, 1.1).fillColor('#d7c28a').fill()
  }

  const starOuter = outerR - 14
  const starInner = outerR - 19
  const starPoints: [number, number][] = []
  for (let i = 0; i < 8; i++) {
    const outerAngle = (i / 8) * Math.PI * 2 - Math.PI / 2
    const innerAngle = outerAngle + Math.PI / 8
    starPoints.push([cx + starOuter * Math.cos(outerAngle), cy + starOuter * Math.sin(outerAngle)])
    starPoints.push([cx + starInner * Math.cos(innerAngle), cy + starInner * Math.sin(innerAngle)])
  }
  doc.polygon(...starPoints).fillColor(accentColor).fillOpacity(0.85).fill()

  drawSealArcText(doc, cx, cy, outerR - 11.5, 'AUTHENTIC', accentColor, 'top')
  drawSealArcText(doc, cx, cy, outerR - 11.5, 'VERIFIED', accentColor, 'bottom')

  doc.restore()
}

const PAGE_W = 842
const PAGE_H = 595
const FRAME_PAD = 40
const FOOTER_BLOCK_H = 130
const MAX_SIG_BOX_Y = PAGE_H - FRAME_PAD - FOOTER_BLOCK_H

function renderCertificatePage(doc: PDFKit.PDFDocument, template: TemplatePayload, recipient: Record<string, unknown>) {
  const name = value(recipient, template.primaryField) || 'Penerima'
  const accentColor = template.accent || '#0f766e'
  const certNumber = getCertificateNumber(template, recipient)

  doc.rect(0, 0, PAGE_W, PAGE_H).fill('#ffffff')

  doc.rect(20, 20, 802, 555).lineWidth(4).stroke(accentColor)
  doc.rect(28, 28, 786, 539).lineWidth(1).stroke('#d7c28a')

  const drawCorner = (x: number, y: number, dx: number, dy: number) => {
    doc.save()
    doc.lineWidth(2).strokeColor(accentColor)
    doc.moveTo(x, y + dy * 22).lineTo(x, y).lineTo(x + dx * 22, y).stroke()
    doc.circle(x + dx * 8, y + dy * 8, 2.5).fillColor('#d7c28a').fill()
    doc.restore()
  }
  drawCorner(34, 34, 1, 1)
  drawCorner(808, 34, -1, 1)
  drawCorner(34, 561, 1, -1)
  drawCorner(808, 561, -1, -1)

  const logoMaxW = 120
  const logoMaxH = 72
  const logoY = 48
  const isLogo2Active = Boolean(template.enableLogo2)

  if (isLogo2Active && template.logo2 && template.logo2.includes('base64,')) {
    try {
      const base64 = template.logo2.substring(template.logo2.indexOf('base64,') + 7)
      if (base64) {
        const logo2Buffer = Buffer.from(base64, 'base64')
        doc.image(logo2Buffer, 48, logoY, { fit: [logoMaxW, logoMaxH], valign: 'center' })
      }
    } catch (err) {
      console.error('Failed to render logo2 in PDF:', err)
    }
  }

  if (template.logo && template.logo.includes('base64,')) {
    try {
      const base64 = template.logo.substring(template.logo.indexOf('base64,') + 7)
      if (base64) {
        const logo1Buffer = Buffer.from(base64, 'base64')
        doc.image(logo1Buffer, PAGE_W - 48 - logoMaxW, logoY, { fit: [logoMaxW, logoMaxH], align: 'right', valign: 'center' })
      }
    } catch (err) {
      console.error('Failed to render logo1 in PDF:', err)
    }
  }

  if (!template.logo && (!isLogo2Active || !template.logo2)) {
    doc.save()
    doc.circle(421, 68, 18).lineWidth(1.5).strokeColor('#d7c28a').stroke()
    doc.circle(421, 68, 14).lineWidth(1).strokeColor(accentColor).stroke()
    doc.fillColor(accentColor).fontSize(15).font('Helvetica-Bold').text('✦', 415, 61, { lineBreak: false, width: 20 })
    doc.restore()
  }

  const fullContentW = PAGE_W - 2 * 90   // 662 (content area dari x=90 sampai x=752)
  const contentX = 90

  const bodyText = interpolate(template.body, recipient, template)
  const bodyFontSize = 13
  const bodyWidth = fullContentW
  const bodyLineGap = 6

  doc.font('Helvetica').fontSize(bodyFontSize)
  const bodyHeight = doc.heightOfString(bodyText, {
    width: bodyWidth,
    align: 'center',
    lineGap: bodyLineGap
  })

  const minTopMargin = 118
  const maxBodyStartY = MAX_SIG_BOX_Y - 28 - bodyHeight - 12
  let currentY = Math.min(minTopMargin, maxBodyStartY < minTopMargin ? maxBodyStartY : minTopMargin)

  if (maxBodyStartY < minTopMargin) currentY = Math.max(FRAME_PAD + 30, maxBodyStartY)

  doc.fillColor(accentColor)
    .fontSize(12)
    .font('Helvetica-Bold')
    .text(template.organization.toUpperCase(), contentX, currentY, { align: 'center', width: fullContentW, characterSpacing: 2 })

  currentY += 30

  doc.fillColor('#172327')
    .fontSize(36)
    .font('Times-Bold')
    .text(template.title, contentX, currentY, { align: 'center', width: fullContentW })

  currentY += 50

  if (template.subtitle && template.subtitle.trim()) {
    doc.fillColor('#4b5563')
      .fontSize(13)
      .font('Helvetica-Oblique')
      .text(template.subtitle.trim(), contentX, currentY, { align: 'center', width: fullContentW })
    currentY += 30
  }

  const typeText = `SERTIFIKAT ${template.type.toUpperCase()}`
  doc.fillColor('#606f7b')
    .fontSize(10.5)
    .font('Helvetica-Bold')
    .text(typeText, contentX, currentY, { align: 'center', width: fullContentW, characterSpacing: 2.5 })

  currentY += 34

  doc.fillColor('#4b5563')
    .fontSize(12.5)
    .font('Helvetica')
    .text('Diberikan dengan penuh kehormatan kepada', contentX, currentY, { align: 'center', width: fullContentW })

  currentY += 32

  doc.fillColor(accentColor)
    .fontSize(36)
    .font('Times-BoldItalic')
    .text(name, contentX, currentY, { align: 'center', width: fullContentW })

  currentY += 54

  const dividerW = 180
  const centerX = PAGE_W / 2
  doc.moveTo(centerX - dividerW - 11, currentY).lineTo(centerX - 11, currentY).lineWidth(1.2).strokeColor('#d7c28a').stroke()
  doc.polygon([centerX, currentY - 4.5], [centerX + 6, currentY], [centerX, currentY + 4.5], [centerX - 6, currentY]).fillColor(accentColor).fill()
  doc.moveTo(centerX + 11, currentY).lineTo(centerX + dividerW + 11, currentY).lineWidth(1.2).strokeColor('#d7c28a').stroke()

  currentY += 26

  const bodyStartY = currentY
  const bodyEndY = bodyStartY + bodyHeight + 12
  const sigBoxY = Math.min(MAX_SIG_BOX_Y, Math.max(bodyEndY + 28, 436))

  doc.fillColor('#334155')
    .fontSize(bodyFontSize)
    .font('Helvetica')
    .text(bodyText, contentX, bodyStartY, {
      align: 'center',
      width: fullContentW,
      lineGap: bodyLineGap
    })

  const isSig2Active = Boolean(template.enableSig2)

  if (isSig2Active) {
    const sealCY = sigBoxY + 38
    drawOfficialSeal(doc, 421, sealCY, accentColor, 'md')

    doc.save()
    doc.fillColor('#64748b')
      .fontSize(8.5)
      .font('Helvetica-Bold')
      .text('NO. SERTIFIKAT', 280, sealCY + 36, { width: 282, align: 'center', characterSpacing: 1.5 })
    doc.fillColor('#1e293b')
      .fontSize(10)
      .font('Helvetica')
      .text(certNumber, 280, sealCY + 50, { width: 282, align: 'center', characterSpacing: 1.2 })
    doc.restore()

    const sig2Label = template.signatoryTitle2 || 'Mengetahui'
    const sig2Name = template.signatory2 || template.organization

    if (template.signatureImage2 && template.signatureImage2.includes('base64,')) {
      try {
        const sig2Base64 = template.signatureImage2.substring(template.signatureImage2.indexOf('base64,') + 7)
        if (sig2Base64) {
          const sig2Buffer = Buffer.from(sig2Base64, 'base64')
          doc.image(sig2Buffer, 60, sigBoxY, { fit: [220, 86], align: 'center', valign: 'center' })
        }
      } catch (err) {
        console.error('Failed to render signature 2 in PDF:', err)
      }
    }

    doc.fillColor('#172327').font('Times-Bold').fontSize(16).text(sig2Name, 55, sigBoxY + 88, { width: 210, align: 'center' })
    doc.moveTo(75, sigBoxY + 110).lineTo(245, sigBoxY + 110).lineWidth(1).strokeColor('#cbd5e1').stroke()
    doc.fillColor('#64748b').font('Helvetica-Bold').fontSize(9.5).text(sig2Label, 55, sigBoxY + 115, { width: 210, align: 'center' })
  } else {
    const sealCY = sigBoxY + 38
    drawOfficialSeal(doc, 155, sealCY, accentColor, 'lg')

    doc.save()
    doc.fillColor('#64748b')
      .fontSize(8.5)
      .font('Helvetica-Bold')
      .text('NO. SERTIFIKAT', 250, sealCY + 32, { width: 342, align: 'center', characterSpacing: 1.5 })
    doc.fillColor('#1e293b')
      .fontSize(10)
      .font('Helvetica')
      .text(certNumber, 250, sealCY + 46, { width: 342, align: 'center', characterSpacing: 1.2 })
    doc.restore()
  }

  const sig1Label = template.signatoryTitle1 || 'Ditetapkan secara resmi oleh'
  const sig1Name = template.signatory || 'Penandatangan'

  if (template.signatureImage && template.signatureImage.includes('base64,')) {
    try {
      const sigBase64 = template.signatureImage.substring(template.signatureImage.indexOf('base64,') + 7)
      if (sigBase64) {
        const sigBuffer = Buffer.from(sigBase64, 'base64')
        doc.image(sigBuffer, 580, sigBoxY, { fit: [220, 86], align: 'center', valign: 'center' })
      }
    } catch (err) {
      console.error('Failed to render signature 1 in PDF:', err)
    }
  }

  doc.fillColor('#172327').font('Times-Bold').fontSize(16).text(sig1Name, 577, sigBoxY + 88, { width: 210, align: 'center' })
  doc.moveTo(597, sigBoxY + 110).lineTo(767, sigBoxY + 110).lineWidth(1).strokeColor('#cbd5e1').stroke()
  doc.fillColor('#64748b').font('Helvetica-Bold').fontSize(9.5).text(sig1Label, 577, sigBoxY + 115, { width: 210, align: 'center' })
}

function generateSinglePDFBuffer(template: TemplatePayload, recipient: Record<string, unknown>): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: 'A4', layout: 'landscape', margin: 0, autoFirstPage: true })
    const chunks: Buffer[] = []
    doc.on('data', (chunk) => chunks.push(chunk))
    doc.on('end', () => resolve(Buffer.concat(chunks)))
    doc.on('error', reject)
    renderCertificatePage(doc, template, recipient)
    doc.end()
  })
}

// Single Certificate PDF Download
app.post('/api/certificates/pdf', async (req, res) => {
  const { template, recipient } = req.body as { template: TemplatePayload; recipient: Record<string, unknown> }
  if (!template || !recipient) return res.status(400).json({ message: 'Template dan penerima wajib ada.' })
  try {
    const name = value(recipient, template.primaryField) || 'Penerima'
    const pdfBuffer = await generateSinglePDFBuffer(template, recipient)
    const filename = `sertifikat-${name.toLowerCase().replace(/[^a-z0-9]+/gi, '-')}.pdf`
    res.setHeader('Content-Type', 'application/pdf')
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`)
    res.setHeader('Content-Length', pdfBuffer.length)
    res.end(pdfBuffer)
  } catch (err) {
    console.error('Error generating PDF:', err)
    if (!res.headersSent) res.status(500).json({ message: 'Gagal membuat file PDF' })
  }
})

// Batch PDF Download: All certificates combined into one multi-page PDF (streaming for quick jobs)
app.post('/api/certificates/batch-pdf', async (req, res) => {
  const { template, recipients } = req.body as { template: TemplatePayload; recipients: Record<string, unknown>[] }
  if (!template || !recipients || !recipients.length) {
    return res.status(400).json({ message: 'Template dan daftar penerima wajib ada.' })
  }
  try {
    const filename = `semua-sertifikat-${(template.name || 'sertifikat').toLowerCase().replace(/[^a-z0-9]+/gi, '-')}.pdf`
    res.status(200)
    res.setHeader('Content-Type', 'application/pdf')
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`)
    res.setHeader('Transfer-Encoding', 'chunked')
    res.flushHeaders?.()

    const doc = new PDFDocument({ size: 'A4', layout: 'landscape', margin: 0, autoFirstPage: false, bufferPages: false })
    doc.pipe(res, { end: true })

    const total = recipients.length
    const chunkSize = 40
    for (let start = 0; start < total; start += chunkSize) {
      const end = Math.min(start + chunkSize, total)
      for (let i = start; i < end; i++) {
        doc.addPage({ size: 'A4', layout: 'landscape', margin: 0 })
        renderCertificatePage(doc, template, recipients[i]!)
      }
      await new Promise<void>((resolve) => setImmediate(resolve))
    }
    doc.end()
  } catch (err) {
    console.error('Error generating batch PDF:', err)
    if (!res.headersSent) res.status(500).json({ message: 'Gagal membuat file batch PDF' })
  }
})

// ============== ASYNC JOB SYSTEM (BYPASS CLOUDFLARE 100s TIMEOUT) ==============
type JobKind = 'zip' | 'batch-pdf'
type JobStatus = 'queued' | 'processing' | 'done' | 'error'
interface Job {
  id: string
  kind: JobKind
  status: JobStatus
  progress: number        // 0..100
  total: number
  done: number
  filename?: string
  filePath?: string
  sizeBytes?: number
  errorMessage?: string
  createdAt: number
  expiresAt: number
}

const jobs = new Map<string, Job>()
const TMP_DIR = fs.existsSync('/tmp') ? '/tmp' : os.tmpdir()
const JOB_TTL_MS = 60 * 60 * 1000      // 1 jam
const JOB_GC_MS = 10 * 60 * 1000        // GC setiap 10 menit

if (!fs.existsSync(TMP_DIR)) fs.mkdirSync(TMP_DIR, { recursive: true })

setInterval(() => {
  const now = Date.now()
  for (const [id, job] of jobs) {
    if (now >= job.expiresAt) {
      if (job.filePath && fs.existsSync(job.filePath)) {
        try { fs.unlinkSync(job.filePath) } catch { /* ignore */ }
      }
      jobs.delete(id)
    }
  }
}, JOB_GC_MS).unref()

function createJob(kind: JobKind, total: number, filename: string): Job {
  const id = randomUUID()
  const job: Job = {
    id,
    kind,
    status: 'queued',
    progress: 0,
    total,
    done: 0,
    filename,
    createdAt: Date.now(),
    expiresAt: Date.now() + JOB_TTL_MS,
  }
  jobs.set(id, job)
  return job
}

async function buildZipJob(template: TemplatePayload, recipients: Record<string, unknown>[], job: Job) {
  try {
    job.status = 'processing'
    const outPath = path.join(TMP_DIR, `certzip-${job.id}.zip`)
    job.filePath = outPath
    const usedFilenames = new Map<string, number>()
    const total = recipients.length

    const archive = new ZipArchive({ zlib: { level: 3 }, forceZip64: true })
    const outStream = fs.createWriteStream(outPath)
    archive.pipe(outStream)

    let cursor = 0
    const concurrency = Math.min(Math.max(2, Math.ceil(total / 50) + 2), 6)
    let completed = 0

    const worker = async () => {
      while (cursor < total) {
        const i = cursor++
        const r = recipients[i]!
        const recipientName = value(r, template.primaryField) || `Penerima-${i + 1}`
        const safeName = recipientName.toLowerCase().replace(/[^a-z0-9]+/gi, '-')
        const count = (usedFilenames.get(safeName) || 0) + 1
        usedFilenames.set(safeName, count)
        const entryName = count > 1
          ? `sertifikat-${safeName}-${count}.pdf`
          : `sertifikat-${safeName}.pdf`
        try {
          const pdfBuffer = await generateSinglePDFBuffer(template, r)
          archive.append(pdfBuffer, { name: entryName, date: new Date() })
        } catch (err) {
          console.error(`Skip PDF ${recipientName}:`, err)
        }
        completed++
        job.done = completed
        job.progress = Math.min(99, Math.round((completed / total) * 100))
        if (completed % 10 === 0) await new Promise<void>((r) => setImmediate(r))
      }
    }

    const workers: Promise<void>[] = []
    for (let w = 0; w < concurrency; w++) workers.push(worker())
    await Promise.all(workers)

    await archive.finalize()
    await new Promise<void>((resolve, reject) => {
      outStream.once('finish', resolve)
      outStream.once('error', reject)
    })

    const stat = fs.statSync(outPath)
    job.sizeBytes = stat.size
    job.progress = 100
    job.done = total
    job.status = 'done'
  } catch (err) {
    console.error('ZIP job failed:', err)
    job.status = 'error'
    job.errorMessage = err instanceof Error ? err.message : 'Gagal membuat ZIP'
  }
}

async function buildBatchPdfJob(template: TemplatePayload, recipients: Record<string, unknown>[], job: Job) {
  try {
    job.status = 'processing'
    const outPath = path.join(TMP_DIR, `certpdf-${job.id}.pdf`)
    job.filePath = outPath
    const total = recipients.length
    const outStream = fs.createWriteStream(outPath)
    const doc = new PDFDocument({ size: 'A4', layout: 'landscape', margin: 0, autoFirstPage: false, bufferPages: false })
    doc.pipe(outStream)

    for (let i = 0; i < total; i++) {
      doc.addPage({ size: 'A4', layout: 'landscape', margin: 0 })
      renderCertificatePage(doc, template, recipients[i]!)
      if ((i + 1) % 20 === 0) {
        job.done = i + 1
        job.progress = Math.min(99, Math.round(((i + 1) / total) * 100))
        await new Promise<void>((r) => setImmediate(r))
      }
    }
    doc.end()

    await new Promise<void>((resolve, reject) => {
      outStream.once('finish', resolve)
      outStream.once('error', reject)
    })
    const stat = fs.statSync(outPath)
    job.sizeBytes = stat.size
    job.progress = 100
    job.done = total
    job.status = 'done'
  } catch (err) {
    console.error('Batch-PDF job failed:', err)
    job.status = 'error'
    job.errorMessage = err instanceof Error ? err.message : 'Gagal membuat PDF gabungan'
  }
}

// ASYNC INIT (Fast response < 100ms)
app.post('/api/certificates/zip-init', async (req, res) => {
  const { template, recipients } = req.body as { template: TemplatePayload; recipients: Record<string, unknown>[] }
  if (!template || !recipients || !recipients.length) {
    return res.status(400).json({ message: 'Template dan daftar penerima wajib ada.' })
  }
  const filename = `sertifikat-lengkap-${(template.name || 'sertifikat').toLowerCase().replace(/[^a-z0-9]+/gi, '-')}.zip`
  const job = createJob('zip', recipients.length, filename)
  setImmediate(() => buildZipJob(template, recipients, job))
  res.status(200).json({ jobId: job.id, total: job.total })
})

app.post('/api/certificates/batch-pdf-init', async (req, res) => {
  const { template, recipients } = req.body as { template: TemplatePayload; recipients: Record<string, unknown>[] }
  if (!template || !recipients || !recipients.length) {
    return res.status(400).json({ message: 'Template dan daftar penerima wajib ada.' })
  }
  const filename = `semua-sertifikat-${(template.name || 'sertifikat').toLowerCase().replace(/[^a-z0-9]+/gi, '-')}.pdf`
  const job = createJob('batch-pdf', recipients.length, filename)
  setImmediate(() => buildBatchPdfJob(template, recipients, job))
  res.status(200).json({ jobId: job.id, total: job.total })
})

// Polling status (hit tiap 1-2 detik)
app.get('/api/certificates/job-status/:jobId', (req, res) => {
  const job = jobs.get(req.params.jobId)
  if (!job) return res.status(404).json({ message: 'Job tidak ditemukan / sudah kadaluarsa.' })
  res.status(200).json({
    id: job.id,
    kind: job.kind,
    status: job.status,
    progress: job.progress,
    total: job.total,
    done: job.done,
    filename: job.filename,
    sizeBytes: job.sizeBytes,
    errorMessage: job.errorMessage,
  })
})

// Download hasil job (hanya kalau status done)
app.get('/api/certificates/job-download/:jobId', (req, res) => {
  const job = jobs.get(req.params.jobId)
  if (!job) return res.status(404).json({ message: 'Job tidak ditemukan / sudah kadaluarsa.' })
  if (job.status === 'error') return res.status(500).json({ message: job.errorMessage || 'Job gagal.' })
  if (job.status !== 'done' || !job.filePath) return res.status(409).json({ message: 'Job belum selesai.' })
  if (!fs.existsSync(job.filePath)) return res.status(410).json({ message: 'File sudah dihapus.' })

  const contentType = job.kind === 'zip' ? 'application/zip' : 'application/pdf'
  res.setHeader('Content-Type', contentType)
  res.setHeader('Content-Disposition', `attachment; filename="${job.filename || 'download'}"`)
  res.setHeader('Content-Length', job.sizeBytes || fs.statSync(job.filePath).size)

  const stream = fs.createReadStream(job.filePath)
  stream.pipe(res)
  stream.on('error', () => res.status(500).end())
  res.on('close', () => {
    try { stream.close(); fs.unlinkSync(job.filePath!) } catch { /* ignore */ }
    jobs.delete(job.id)
  })
})

// Legacy ZIP / batch-pdf endpoints -> direct delegate to async for large, stream for small
app.post('/api/certificates/zip', async (req, res) => {
  const { template, recipients } = req.body as { template: TemplatePayload; recipients: Record<string, unknown>[] }
  if (!template || !recipients || !recipients.length) {
    return res.status(400).json({ message: 'Template dan daftar penerima wajib ada.' })
  }
  if (recipients.length <= 30) {
    try {
      const filename = `sertifikat-lengkap-${(template.name || 'sertifikat').toLowerCase().replace(/[^a-z0-9]+/gi, '-')}.zip`
      res.status(200)
      res.setHeader('Content-Type', 'application/zip')
      res.setHeader('Content-Disposition', `attachment; filename="${filename}"`)
      res.setHeader('Transfer-Encoding', 'chunked')
      res.flushHeaders?.()

      const archive = new ZipArchive({ zlib: { level: 3 }, forceZip64: true })
      archive.pipe(res, { end: true })
      const usedFilenames = new Map<string, number>()
      for (let i = 0; i < recipients.length; i++) {
        const r = recipients[i]!
        const recipientName = value(r, template.primaryField) || `Penerima-${i + 1}`
        const safeName = recipientName.toLowerCase().replace(/[^a-z0-9]+/gi, '-')
        const count = (usedFilenames.get(safeName) || 0) + 1
        usedFilenames.set(safeName, count)
        const entryName = count > 1 ? `sertifikat-${safeName}-${count}.pdf` : `sertifikat-${safeName}.pdf`
        try {
          const pdfBuffer = await generateSinglePDFBuffer(template, r)
          archive.append(pdfBuffer, { name: entryName, date: new Date() })
        } catch (err) {
          console.error(`Skip PDF ${recipientName}:`, err)
        }
      }
      await archive.finalize()
    } catch (err) {
      console.error('Error generating ZIP:', err)
      if (!res.headersSent) res.status(500).json({ message: 'Gagal membuat file ZIP' })
    }
    return
  }
  // Large batch: return jobId untuk polling (frontend yang support async akan mengikuti)
  res.status(202).json({
    message: 'Batch besar terdeteksi, gunakan mode async dengan /zip-init.',
    total: recipients.length,
    initEndpoint: '/api/certificates/zip-init',
  })
})

export default app

const fs = require('fs')
const path = require('path')
const XLSX = require('xlsx')

let dataDir = path.join(__dirname, '../../data')
let defaultExcelPath = path.join(dataDir, '2025年业绩.xlsx')
let pendingPath = path.join(dataDir, '.pending_changes.json')

function setDataDir(dir) {
  dataDir = dir
  defaultExcelPath = path.join(dataDir, '2025年业绩.xlsx')
  pendingPath = path.join(dataDir, '.pending_changes.json')
}

// 默认数据结构（字段与 2025年业绩.xlsx 对齐）
const defaultHeaders = ['日期', '姓名', '卡种', '实缴金额', '尾款金额', '顾问', '试课老师', '会籍单', '来源', '备注']

// 当前选中的月份
let currentMonth = null

// === 待处理更改队列（Excel 文件被锁定时暂存） ===

function loadPendingChanges() {
  try {
    if (fs.existsSync(pendingPath)) {
      return JSON.parse(fs.readFileSync(pendingPath, 'utf-8'))
    }
  } catch (e) {
    console.error('读取待处理更改失败:', e)
  }
  return { adds: [], deletes: [] }
}

function savePendingChanges(pending) {
  try {
    if (!fs.existsSync(dataDir)) {
      fs.mkdirSync(dataDir, { recursive: true })
    }
    fs.writeFileSync(pendingPath, JSON.stringify(pending, null, 2), 'utf-8')
  } catch (e) {
    console.error('保存待处理更改失败:', e)
  }
}

function hasPendingChanges() {
  const pending = loadPendingChanges()
  return pending.adds.length > 0 || pending.deletes.length > 0
}

// 尝试将待处理更改写入 Excel
function tryFlushPending() {
  const pending = loadPendingChanges()
  if (pending.adds.length === 0 && pending.deletes.length === 0) {
    return { flushed: true, pendingCount: 0 }
  }

  // 检测文件是否可写
  try {
    const fd = fs.openSync(excelPath, 'r+')
    fs.closeSync(fd)
  } catch (e) {
    return { flushed: false, pendingCount: pending.adds.length + pending.deletes.length }
  }

  // 文件可写，应用待处理更改
  try {
    const workbook = XLSX.readFile(excelPath)

    for (const addData of pending.adds) {
      const sheetName = addData.__pendingSheet || getTargetSheetName(workbook, null)
      appendPerformanceRow(workbook, sheetName, addData)
    }

    for (const delData of pending.deletes) {
      const worksheet = workbook.Sheets[delData.__sheetName]
      if (worksheet) {
        const sheetJson = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' })
        const rowIndex = Number(delData.__rowIndex)
        if (rowIndex > 0 && rowIndex < sheetJson.length) {
          sheetJson.splice(rowIndex, 1)
          workbook.Sheets[delData.__sheetName] = XLSX.utils.aoa_to_sheet(sheetJson)
        }
      }
    }

    XLSX.writeFile(workbook, excelPath)
    savePendingChanges({ adds: [], deletes: [] })
    console.log('待处理更改已同步到 Excel')
    return { flushed: true, pendingCount: 0 }
  } catch (e) {
    console.error('应用待处理更改失败:', e)
    return { flushed: false, pendingCount: pending.adds.length + pending.deletes.length }
  }
}

// 添加操作到待处理队列
function queuePendingAdd(data, sheetName) {
  const pending = loadPendingChanges()
  pending.adds.push({ ...data, __pendingSheet: sheetName })
  savePendingChanges(pending)
}

// 添加删除操作到待处理队列
function queuePendingDelete(record) {
  const pending = loadPendingChanges()
  pending.deletes.push({
    __sheetName: record.__sheetName,
    __rowIndex: record.__rowIndex
  })
  savePendingChanges(pending)
}

function isFileLocked() {
  try {
    const fd = fs.openSync(excelPath, 'r+')
    fs.closeSync(fd)
    return false
  } catch (e) {
    return e.code === 'EBUSY'
  }
}
let excelPath = defaultExcelPath

function setExcelPath(filePath) {
  excelPath = filePath || defaultExcelPath
  return excelPath
}

// 安全写入 Excel 文件，处理文件被锁定的情况
function safeWriteFile(workbook, filePath) {
  try {
    XLSX.writeFile(workbook, filePath)
    return { success: true }
  } catch (error) {
    if (error.code === 'EBUSY') {
      return { success: false, error: '文件被占用，请关闭 Excel 后重试' }
    }
    throw error
  }
}

function getExcelPath() {
  return excelPath
}

function getExcelBindingInfo() {
  return {
    filePath: excelPath,
    fileName: path.basename(excelPath),
    exists: fs.existsSync(excelPath)
  }
}

function formatChineseDate(date = new Date()) {
  return `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}号`
}

function parseDateToTime(value) {
  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value)
    if (parsed) {
      return new Date(parsed.y, parsed.m - 1, parsed.d).getTime()
    }
    return value
  }

  const str = String(value || '').trim()
  if (!str) return 0

  const chineseMatch = str.match(/(\d{4})年(\d{1,2})月(\d{1,2})号/)
  if (chineseMatch) {
    const [, year, month, day] = chineseMatch
    return new Date(Number(year), Number(month) - 1, Number(day)).getTime()
  }

  const normalized = str
    .replace(/年/g, '/')
    .replace(/月/g, '/')
    .replace(/号/g, '')
    .replace(/日/g, '')

  const timestamp = new Date(normalized).getTime()
  return Number.isNaN(timestamp) ? 0 : timestamp
}

function isSameDate(value, date = new Date()) {
  const timestamp = parseDateToTime(value)
  if (!timestamp) return false

  const target = new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime()
  const current = new Date(new Date(timestamp).getFullYear(), new Date(timestamp).getMonth(), new Date(timestamp).getDate()).getTime()
  return current === target
}

function getValue(row, keys, fallback = '') {
  for (const key of keys) {
    const value = row?.[key]
    if (value !== undefined && value !== null && value !== '') {
      return value
    }
  }
  return fallback
}

function getCustomerName(row) {
  return getValue(row, ['姓名', '客户姓名', '客户'], '')
}

function getTeacher(row) {
  return getValue(row, ['顾问', '老师', '销售', 'zzz'], '')
}

function getAmount(row) {
  return getValue(row, ['实缴金额', '金额', '业绩'], 0)
}

function getRemark(row) {
  return getValue(row, ['备注', '说明'], '')
}

function getSyncStatus(row) {
  return getValue(row, ['是否同步'], '是')
}

function normalizePerformanceRow(row = {}) {
  const customerName = getCustomerName(row)
  const teacher = getTeacher(row)
  const amount = getAmount(row)
  const remark = getRemark(row)
  const syncStatus = getSyncStatus(row)

  return {
    ...row,
    '姓名': customerName,
    '客户姓名': customerName,
    '顾问': teacher,
    '老师': teacher,
    '销售': teacher,
    'zzz': teacher,
    '实缴金额': amount,
    '金额': amount,
    '业绩': amount,
    '备注': remark,
    '是否同步': syncStatus
  }
}

function buildPerformanceRow(data = {}) {
  const date = data.date || data['日期'] || formatChineseDate()
  const customerName = data.customerName || data['姓名'] || data['客户姓名'] || data['客户'] || ''
  const teacher = data.teacher || data['顾问'] || data['老师'] || data['销售'] || data['zzz'] || ''
  const cardType = data.cardType || data['卡种'] || data['卡类型'] || ''
  const amount = data.amount ?? data['实缴金额'] ?? data['金额'] ?? data['业绩'] ?? 0
  const tailAmount = data.tailAmount ?? data['尾款金额'] ?? ''
  const trialTeacher = data.trialTeacher || data['试课老师'] || ''
  const membership = data.membership || data['会籍单'] || ''
  const source = data.source || data['来源'] || ''
  const remark = data.remark || data['备注'] || data['说明'] || ''

  return {
    '日期': date,
    '姓名': customerName,
    '客户姓名': customerName,
    '卡种': cardType,
    '实缴金额': amount,
    '金额': amount,
    '业绩': amount,
    '尾款金额': tailAmount,
    '顾问': teacher,
    '老师': teacher,
    '销售': teacher,
    'zzz': teacher,
    '试课老师': trialTeacher,
    '会籍单': membership,
    '来源': source,
    '备注': remark
  }
}

function getSheetRows(worksheet) {
  return XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' })
}

function mapSheetRowsToRecords(sheetName, worksheet) {
  const rows = getSheetRows(worksheet)
  const headers = rows[0] || []

  return rows.slice(1)
    .map((rowValues, index) => {
      const row = {}
      headers.forEach((header, headerIndex) => {
        if (header) {
          row[header] = rowValues[headerIndex]
        }
      })

      const normalized = normalizePerformanceRow(row)
      const hasData = normalized['日期'] || normalized['姓名'] || normalized['卡种'] || normalized['顾问'] || normalized['实缴金额']
      if (!hasData) return null

      return {
        ...normalized,
        __sheetName: sheetName,
        __rowIndex: index + 1
      }
    })
    .filter(Boolean)
}

function sortRecordsByDateDesc(data = []) {
  return [...data].sort((a, b) => parseDateToTime(b['日期']) - parseDateToTime(a['日期']))
}

function getSheetNameList(workbook) {
  const sheetNames = workbook.SheetNames.filter(name => name !== '业绩表' && !/^Sheet\d+$/.test(name))
  return sheetNames.length > 0 ? sheetNames : workbook.SheetNames.filter(name => name !== '业绩表')
}

function getTargetSheetName(workbook, month = null) {
  if (month && workbook.Sheets[month]) {
    return month
  }

  if (currentMonth && workbook.Sheets[currentMonth]) {
    return currentMonth
  }

  if (workbook.Sheets['业绩表']) {
    return '业绩表'
  }

  const sheetNames = getSheetNameList(workbook)
  if (sheetNames.length > 0) {
    return sheetNames[sheetNames.length - 1]
  }

  return '业绩表'
}

function ensureSheet(workbook, sheetName) {
  if (!workbook.Sheets[sheetName]) {
    const worksheet = XLSX.utils.aoa_to_sheet([defaultHeaders])
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName)
  }

  return workbook.Sheets[sheetName]
}

function appendPerformanceRow(workbook, sheetName, data) {
  const worksheet = ensureSheet(workbook, sheetName)
  const rows = getSheetRows(worksheet)
  const headers = rows[0] && rows[0].length > 0 ? rows[0] : defaultHeaders
  const row = buildPerformanceRow(data)

  rows.push(headers.map(header => {
    if (!header) return ''
    return row[header] ?? ''
  }))

  workbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(rows)
}

// 初始化 Excel 文件
function initExcelFile() {
  // 确保数据目录存在
  if (!fs.existsSync(dataDir)) {
    fs.mkdirSync(dataDir, { recursive: true })
  }

  // 如果 Excel 文件不存在，创建一个空的
  if (!fs.existsSync(excelPath)) {
    const workbook = XLSX.utils.book_new()
    const worksheet = XLSX.utils.aoa_to_sheet([defaultHeaders])
    XLSX.utils.book_append_sheet(workbook, worksheet, '业绩表')
    safeWriteFile(workbook, excelPath)
    console.log('已创建新的 2025年业绩.xlsx')
  }
}

// 获取所有月份列表
function getMonthList() {
  try {
    if (!fs.existsSync(excelPath)) {
      return []
    }

    const workbook = XLSX.readFile(excelPath)
    return getSheetNameList(workbook)
  } catch (error) {
    console.error('获取月份列表失败:', error)
    return []
  }
}

// 设置当前月份
function setCurrentMonth(month) {
  currentMonth = month
  return { success: true, month }
}

function serializeRecord(record) {
  const { __sheetName, __rowIndex, ...rest } = record
  const serialized = {}
  for (const key of Object.keys(rest)) {
    const value = rest[key]
    if (typeof value === 'object' && value !== null) {
      serialized[key] = JSON.parse(JSON.stringify(value))
    } else {
      serialized[key] = value
    }
  }
  return serialized
}

function serializeData(data) {
  return data.map(record => serializeRecord(record))
}

// 读取业绩数据（支持按月份读取）
function readPerformanceData(month = null) {
  try {
    if (!fs.existsSync(excelPath)) {
      return { success: false, error: '业绩表文件不存在' }
    }

    // 尝试刷新待处理更改
    const flushResult = tryFlushPending()

    const workbook = XLSX.readFile(excelPath)
    const targetSheets = month && workbook.Sheets[month]
      ? [month]
      : getSheetNameList(workbook)

    const data = sortRecordsByDateDesc(targetSheets.flatMap(sheetName => {
      const worksheet = workbook.Sheets[sheetName]
      if (!worksheet) return []
      return mapSheetRowsToRecords(sheetName, worksheet)
    }))

    // 合并待处理的新增记录
    const pending = loadPendingChanges()
    if (pending.adds.length > 0) {
      const pendingRecords = pending.adds
        .filter(add => !month || add.__pendingSheet === month)
        .map((add, index) => {
          const record = buildPerformanceRow(add)
          record.__sheetName = add.__pendingSheet || targetSheets[0] || '业绩表'
          record.__rowIndex = -(index + 1) // 负数标记为待处理
          record.__pending = true
          return record
        })
      data.unshift(...pendingRecords)
    }

    // 标记待删除的记录
    const deleteKeys = new Set(pending.deletes.map(d => `${d.__sheetName}|${d.__rowIndex}`))

    return {
      success: true,
      data: serializeData(data),
      month: month || '全部',
      pendingCount: pending.adds.length + pending.deletes.length,
      flushed: flushResult.flushed
    }
  } catch (error) {
    console.error('读取 Excel 失败:', error)
    return { success: false, error: error.message }
  }
}

// 新增业绩
function addPerformance(data, month = null) {
  try {
    if (!fs.existsSync(excelPath)) {
      initExcelFile()
    }

    // 先尝试刷新待处理队列
    tryFlushPending()

    const workbook = XLSX.readFile(excelPath)
    const sheetName = getTargetSheetName(workbook, month)

    appendPerformanceRow(workbook, sheetName, data)
    const writeResult = safeWriteFile(workbook, excelPath)

    if (!writeResult.success) {
      // Excel 被锁定，暂存到待处理队列
      if (writeResult.error && writeResult.error.includes('文件被占用')) {
        queuePendingAdd(data, sheetName)
        return { success: true, pending: true, message: 'Excel 文件被占用，业绩已暂存，关闭 Excel 后将自动写入' }
      }
      return writeResult
    }

    return { success: true }
  } catch (error) {
    console.error('添加业绩失败:', error)
    return { success: false, error: error.message }
  }
}

// 删除业绩
function deletePerformance(record) {
  try {
    if (!record || !record.__sheetName || record.__rowIndex === undefined) {
      return { success: false, error: '缺少要删除的行信息' }
    }

    const workbook = XLSX.readFile(excelPath)
    const worksheet = workbook.Sheets[record.__sheetName]
    if (!worksheet) {
      return { success: false, error: '对应月份不存在' }
    }

    const sheetJson = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' })
    const rowIndex = Number(record.__rowIndex)
    if (rowIndex <= 0 || rowIndex >= sheetJson.length) {
      return { success: false, error: '对应行不存在或已变化' }
    }

    sheetJson.splice(rowIndex, 1)

    const newWorksheet = XLSX.utils.aoa_to_sheet(sheetJson)
    workbook.Sheets[record.__sheetName] = newWorksheet
    const writeResult = safeWriteFile(workbook, excelPath)

    if (!writeResult.success) {
      // Excel 被锁定，暂存到待处理队列
      if (writeResult.error && writeResult.error.includes('文件被占用')) {
        queuePendingDelete(record)
        return { success: true, pending: true, message: 'Excel 文件被占用，删除已暂存，关闭 Excel 后将自动同步' }
      }
      return writeResult
    }

    return { success: true }
  } catch (error) {
    console.error('删除业绩失败:', error)
    return { success: false, error: error.message }
  }
}

// 获取统计数据
function getStats(month = null, filters = {}) {
  try {
    const result = readPerformanceData(month)
    const monthLabel = month || '全部月份'
    if (!result.success) {
      return { total: 0, byTeacher: {}, count: 0, todayTotal: 0, todayByTeacher: {}, todayCount: 0, pendingCount: 0, monthLabel }
    }

    let data = result.data || []

    // 应用筛选
    if (filters.teacher) {
      data = data.filter(item => item['顾问'] === filters.teacher || item['老师'] === filters.teacher)
    }
    if (filters.cardType) {
      data = data.filter(item => item['卡种'] === filters.cardType)
    }

    const total = data.reduce((sum, item) => sum + (parseFloat(item['实缴金额']) || 0), 0)

    const byTeacher = {}
    data.forEach(item => {
      const teacher = item['顾问'] || item['老师']
      const amount = parseFloat(item['实缴金额']) || 0
      if (teacher) {
        byTeacher[teacher] = (byTeacher[teacher] || 0) + amount
      }
    })

    // 今日统计
    const today = new Date()
    const todayData = data.filter(item => isSameDate(item['日期'], today))
    const todayTotal = todayData.reduce((sum, item) => sum + (parseFloat(item['实缴金额']) || 0), 0)
    const todayByTeacher = {}
    todayData.forEach(item => {
      const teacher = item['顾问'] || item['老师']
      const amount = parseFloat(item['实缴金额']) || 0
      if (teacher) {
        todayByTeacher[teacher] = (todayByTeacher[teacher] || 0) + amount
      }
    })

    return { total, byTeacher, count: data.length, todayTotal, todayByTeacher, todayCount: todayData.length, pendingCount: result.pendingCount || 0, monthLabel }
  } catch (error) {
    console.error('统计失败:', error)
    return { total: 0, byTeacher: {}, count: 0, todayTotal: 0, todayByTeacher: {}, todayCount: 0, pendingCount: 0, monthLabel: month || '全部月份' }
  }
}

// 获取未同步的数据
function getUnsyncedData(month = null) {
  try {
    const result = readPerformanceData(month)
    if (!result.success) return []

    const unsynced = result.data.filter(item =>
      item['是否同步'] === '否' || item['是否同步'] === false
    )

    return unsynced
  } catch (error) {
    console.error('获取未同步数据失败:', error)
    return []
  }
}

// 标记为已同步
function markSynced(indices, month = null) {
  try {
    const workbook = XLSX.readFile(excelPath)
    const sheetName = getTargetSheetName(workbook, month)
    const worksheet = workbook.Sheets[sheetName]
    const sheetJson = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' })
    const headers = sheetJson[0] || []
    const syncColumnIndex = headers.indexOf('是否同步')

    if (syncColumnIndex === -1) {
      return { success: true }
    }

    indices.forEach(index => {
      const rowIndex = index + 1
      if (rowIndex < sheetJson.length) {
        while (sheetJson[rowIndex].length <= syncColumnIndex) {
          sheetJson[rowIndex].push('')
        }
        sheetJson[rowIndex][syncColumnIndex] = '是'
      }
    })

    workbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(sheetJson)
    const writeResult = safeWriteFile(workbook, excelPath)
    if (!writeResult.success) return writeResult

    return { success: true }
  } catch (error) {
    console.error('标记同步失败:', error)
    return { success: false, error: error.message }
  }
}

// 获取老师列表
function getTeacherList(month = null) {
  try {
    const result = readPerformanceData(month)
    if (!result.success) return []

    const teachers = new Set(result.data
      .map(item => item['顾问'] || item['老师'])
      .filter(Boolean)
    )

    return Array.from(teachers)
  } catch (error) {
    return []
  }
}

// 获取卡种列表
function getCardTypes(month = null) {
  try {
    const result = readPerformanceData(month)
    if (!result.success) return []

    const cardTypes = new Set(result.data
      .map(item => item['卡种'])
      .filter(Boolean)
    )

    return Array.from(cardTypes)
  } catch (error) {
    return []
  }
}

// 导入业绩表
async function importPerformance(filePath) {
  try {
    const workbook = XLSX.readFile(filePath)
    const mainWorkbook = XLSX.utils.book_new()

    // 复制所有 sheet 到新文件
    let totalRecords = 0

    workbook.SheetNames.forEach(sheetName => {
      const worksheet = workbook.Sheets[sheetName]
      const data = XLSX.utils.sheet_to_json(worksheet)

      // 统一转换为 2025年业绩.xlsx 字段：客户姓名=姓名，老师=顾问，金额=实缴金额
      const normalizedRows = data
        .map(row => buildPerformanceRow(row))
        .filter(row => row['日期'] || row['姓名'] || row['卡种'] || row['顾问'] || row['实缴金额'])

      const rows = [defaultHeaders, ...normalizedRows.map(row => defaultHeaders.map(header => row[header] ?? ''))]
      const newWorksheet = XLSX.utils.aoa_to_sheet(rows)
      XLSX.utils.book_append_sheet(mainWorkbook, newWorksheet, sheetName)
      totalRecords += normalizedRows.length
    })

    // 添加默认的业绩表 sheet（如果没有）
    if (!mainWorkbook.SheetNames.includes('业绩表')) {
      const emptySheet = XLSX.utils.aoa_to_sheet([defaultHeaders])
      XLSX.utils.book_append_sheet(mainWorkbook, emptySheet, '业绩表')
    }

    // 保存到数据目录
    const writeResult = safeWriteFile(mainWorkbook, excelPath)
    if (!writeResult.success) return writeResult

    return { success: true, count: totalRecords, months: workbook.SheetNames }
  } catch (error) {
    console.error('导入失败:', error)
    return { success: false, error: error.message }
  }
}

// 导出业绩表
function exportPerformance(outputPath, year, month) {
  try {
    const result = readPerformanceData(month || null)
    if (!result.success) {
      return { success: false, error: '读取数据失败' }
    }

    const workbook = XLSX.utils.book_new()
    let data = result.data || []

    // 筛选年份
    if (year) {
      data = data.filter(item => {
        const dateStr = String(item['日期'] || '')
        return dateStr.includes(String(year))
      })
    }

    // 筛选月份
    if (month) {
      const normalizedMonth = Number(month)
      data = data.filter(item => {
        const dateStr = String(item['日期'] || '')
        return dateStr.includes(`${year}-${String(month).padStart(2, '0')}`) ||
               dateStr.includes(`${year}/${String(month).padStart(2, '0')}`) ||
               dateStr.includes(`${year}年${normalizedMonth}月`) ||
               dateStr.includes(`${year}年${String(month).padStart(2, '0')}月`)
      })
    }

    const exportData = data.map(item => ({
      '日期': item['日期'] || '',
      '姓名': item['姓名'] || item['客户姓名'] || '',
      '卡种': item['卡种'] || '',
      '实缴金额': item['实缴金额'] ?? item['金额'] ?? 0,
      '尾款金额': item['尾款金额'] || '',
      '顾问': item['顾问'] || item['老师'] || '',
      '试课老师': item['试课老师'] || '',
      '会籍单': item['会籍单'] || '',
      '来源': item['来源'] || '',
      '备注': item['备注'] || ''
    }))

    const worksheet = XLSX.utils.json_to_sheet(exportData, { header: defaultHeaders })
    const sheetName = month ? `${year}-${month.toString().padStart(2, '0')}` : `${year}年`
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName)
    XLSX.writeFile(workbook, outputPath)

    return { success: true, count: data.length }
  } catch (error) {
    console.error('导出失败:', error)
    return { success: false, error: error.message }
  }
}

// 添加月份（新建 sheet，以模板月份的表头为字段）
function addMonth(newSheetName, templateSheetName) {
  try {
    if (!fs.existsSync(excelPath)) {
      return { success: false, error: '业绩表文件不存在' }
    }

    const workbook = XLSX.readFile(excelPath)

    if (workbook.Sheets[newSheetName]) {
      return { success: false, error: `月份"${newSheetName}"已存在` }
    }

    // 获取模板表头
    let headers = defaultHeaders
    const sheetNames = getSheetNameList(workbook)

    if (templateSheetName && workbook.Sheets[templateSheetName]) {
      const rows = getSheetRows(workbook.Sheets[templateSheetName])
      if (rows.length > 0) {
        headers = rows[0].filter(h => h)
      }
    } else if (sheetNames.length > 0) {
      // 使用最后一个月份 sheet 的表头作为模板
      const latestName = sheetNames[sheetNames.length - 1]
      const rows = getSheetRows(workbook.Sheets[latestName])
      if (rows.length > 0) {
        headers = rows[0].filter(h => h)
      }
    }

    const newWorksheet = XLSX.utils.aoa_to_sheet([headers])
    XLSX.utils.book_append_sheet(workbook, newWorksheet, newSheetName)

    const writeResult = safeWriteFile(workbook, excelPath)
    if (!writeResult.success) return writeResult

    return { success: true, month: newSheetName }
  } catch (error) {
    console.error('添加月份失败:', error)
    return { success: false, error: error.message }
  }
}

// 删除月份
function deleteMonth(sheetName) {
  try {
    if (!fs.existsSync(excelPath)) {
      return { success: false, error: '业绩表文件不存在' }
    }

    const workbook = XLSX.readFile(excelPath)

    if (!workbook.Sheets[sheetName]) {
      return { success: false, error: `月份"${sheetName}"不存在` }
    }

    const sheetNames = getSheetNameList(workbook)
    if (sheetNames.length <= 1) {
      return { success: false, error: '至少保留一个月份' }
    }

    // 从 SheetNames 中移除
    const idx = workbook.SheetNames.indexOf(sheetName)
    if (idx >= 0) {
      workbook.SheetNames.splice(idx, 1)
    }
    delete workbook.Sheets[sheetName]

    const writeResult = safeWriteFile(workbook, excelPath)
    if (!writeResult.success) return writeResult

    return { success: true }
  } catch (error) {
    console.error('删除月份失败:', error)
    return { success: false, error: error.message }
  }
}

module.exports = {
  initExcelFile,
  readPerformanceData,
  addPerformance,
  deletePerformance,
  getStats,
  getUnsyncedData,
  markSynced,
  getTeacherList,
  getCardTypes,
  importPerformance,
  exportPerformance,
  addMonth,
  deleteMonth,
  hasPendingChanges,
  tryFlushPending,
  getMonthList,
  setCurrentMonth,
  setDataDir,
  setExcelPath,
  getExcelPath,
  getExcelBindingInfo,
  formatChineseDate,
  parseDateToTime,
  isSameDate,
  defaultExcelPath
}

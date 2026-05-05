const { app, BrowserWindow, ipcMain, dialog } = require('electron')
const fs = require('fs')
const path = require('path')
const Store = require('electron-store')

// 导入模块
const excelHandler = require('./excel')
const pushHandler = require('./push')
const scheduleHandler = require('./schedule')

// 初始化存储
const store = new Store({
  name: 'config',
  defaults: {
    wechat: {
      webhookUrl: '',
      celebrationTemplate: '',
      selectedTeachers: []
    },
    schedule: {
      enabled: false,
      dailyTime: '21:00',
      intervalEnabled: false,
      intervalMinutes: 60
    },
    excel: {
      boundFilePath: excelHandler.defaultExcelPath
    }
  }
})

let mainWindow
let excelFileWatcher = null
let excelFileWatchTimer = null

function notifyExcelFileChanged(reason = 'updated') {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('excel-file-changed', {
      reason,
      binding: excelHandler.getExcelBindingInfo()
    })
  }
}

function stopExcelFileWatcher() {
  if (excelFileWatcher && excelFileWatcher.filePath) {
    fs.unwatchFile(excelFileWatcher.filePath)
    excelFileWatcher = null
  }
  if (excelFileWatchTimer) {
    clearTimeout(excelFileWatchTimer)
    excelFileWatchTimer = null
  }
}

function watchExcelFile() {
  stopExcelFileWatcher()
  const { filePath, exists } = excelHandler.getExcelBindingInfo()
  if (!filePath || !exists) return

  try {
    fs.watchFile(filePath, { interval: 1000 }, (curr, prev) => {
      if (curr.mtime !== prev.mtime) {
        if (excelFileWatchTimer) {
          clearTimeout(excelFileWatchTimer)
        }
        excelFileWatchTimer = setTimeout(() => {
          notifyExcelFileChanged('external-update')
        }, 300)
      }
    })
    excelFileWatcher = { filePath }
  } catch (error) {
    console.error('监听 Excel 文件失败:', error)
  }
}

function bindExcelFile(filePath, persist = true) {
  const nextPath = filePath || excelHandler.defaultExcelPath
  excelHandler.setExcelPath(nextPath)
  excelHandler.initExcelFile()
  watchExcelFile()

  if (persist) {
    store.set('excel.boundFilePath', nextPath)
  }

  return excelHandler.getExcelBindingInfo()
}

function migratePushConfig() {
  const wechat = store.get('wechat')
  const push = store.get('push') || {}

  // 迁移通用模板和老师筛选到 push 根级别
  if (!push.celebrationTemplate && (wechat?.celebrationTemplate || push.wecomBot?.celebrationTemplate)) {
    push.celebrationTemplate = push.wecomBot?.celebrationTemplate || wechat?.celebrationTemplate || ''
  }
  if (!push.selectedTeachers || push.selectedTeachers.length === 0) {
    const teachers = push.wecomBot?.selectedTeachers || wechat?.selectedTeachers || []
    if (teachers.length > 0) {
      push.selectedTeachers = teachers
    }
  }

  // 迁移 wecomBot 配置
  if (wechat?.webhookUrl && !push.wecomBot?.webhookUrl) {
    console.log('迁移企微配置到多通道推送模块...')
    push.wecomBot = {
      enabled: true,
      webhookUrl: wechat.webhookUrl
    }
  }

  // 确保 wxpusher 默认值存在
  if (!push.wxpusher) {
    push.wxpusher = {
      enabled: false,
      appToken: '',
      topicIds: []
    }
  }

  // 确保 dingtalk 默认值存在
  if (!push.dingtalk) {
    push.dingtalk = {
      enabled: false,
      webhookUrl: '',
      secret: ''
    }
  }

  store.set('push', push)
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    show: false,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      webSecurity: false,
      sandbox: false
    }
  })

  // Win11 兼容：窗口准备好后再显示，避免渲染崩溃导致白屏
  mainWindow.once('ready-to-show', () => {
    mainWindow.show()
  })

  // 开发模式 - 检测是否在 vite 开发服务器环境
  const isDev = process.argv.includes('--dev') || process.env.NODE_ENV === 'development'

  if (isDev) {
    const devUrl = process.env.VITE_DEV_SERVER_URL || 'http://localhost:5173'
    mainWindow.loadURL(devUrl)
    mainWindow.webContents.openDevTools()
  } else {
    // 生产模式
    mainWindow.loadFile(path.join(__dirname, '../../dist/index.html'))
  }

  // 渲染进程错误日志（生产环境调试用）
  mainWindow.webContents.on('console-message', (event, level, message) => {
    console.log(`[RENDERER] ${message}`)
  })

  mainWindow.on('closed', () => {
    mainWindow = null
  })
}

// Win11 兼容：禁用 GPU 加速避免显卡驱动不兼容
app.commandLine.appendSwitch('disable-gpu')
app.commandLine.appendSwitch('disable-software-rasterizer')

app.whenReady().then(() => {
  try {
    // IPC 必须在 createWindow 之前注册，否则渲染进程 invoke 时找不到处理器
    setupIpcHandlers()

    // 设置数据目录（生产环境使用 userData，避免写入只读 asar）
    excelHandler.setDataDir(app.getPath('userData'))

    // 从配置读取绑定路径并初始化
    const boundPath = store.get('excel.boundFilePath') || excelHandler.defaultExcelPath
    bindExcelFile(boundPath, false)

    // 配置迁移：将旧 wechat.* 配置迁移到 push.wecomBot.*
    migratePushConfig()

    // 初始化定时任务
    scheduleHandler.initSchedule(pushHandler)

    createWindow()
  } catch (e) {
    console.error('启动失败:', e)
    // 确保即使出错也能显示窗口
    createWindow()
  }
})

function setupIpcHandlers() {
  ipcMain.handle('get-performance-data', async () => {
    try {
      return JSON.parse(JSON.stringify(excelHandler.readPerformanceData()))
    } catch (e) {
      return { success: false, error: e.message }
    }
  })

  // 新增业绩
  ipcMain.handle('add-performance', async (event, data, month) => {
    try {
      return JSON.parse(JSON.stringify(excelHandler.addPerformance(data, month)))
    } catch (e) {
      return { success: false, error: e.message }
    }
  })

  // 删除业绩
  ipcMain.handle('delete-performance', async (event, record) => {
    try {
      return JSON.parse(JSON.stringify(excelHandler.deletePerformance(record)))
    } catch (e) {
      return { success: false, error: e.message }
    }
  })

  // 获取 Excel 绑定信息
  ipcMain.handle('get-excel-binding-info', async () => {
    try {
      return JSON.parse(JSON.stringify(excelHandler.getExcelBindingInfo()))
    } catch (e) {
      return { filePath: '', fileName: '', exists: false }
    }
  })

  // 绑定 Excel 文件
  ipcMain.handle('bind-excel-file', async (event, filePath) => {
    try {
      const binding = bindExcelFile(filePath, true)
      notifyExcelFileChanged('rebind')
      return { success: true, binding }
    } catch (error) {
      console.error('绑定 Excel 文件失败:', error)
      return { success: false, error: error.message }
    }
  })

  // 获取统计
  ipcMain.handle('get-stats', async (event, filters) => {
    try {
      return JSON.parse(JSON.stringify(excelHandler.getStats(null, filters)))
    } catch (e) {
      return { total: 0, byTeacher: {}, count: 0 }
    }
  })

  // 获取配置
  ipcMain.handle('get-config', async () => {
    try {
      return JSON.parse(JSON.stringify(store.store))
    } catch (e) {
      return {}
    }
  })

  // 保存配置
  ipcMain.handle('save-config', async (event, config) => {
    try {
      console.log('保存配置:', config)
      store.set(config)
      // 更新定时任务
      scheduleHandler.updateSchedule(config, pushHandler)
      console.log('配置保存成功')
      return { success: true }
    } catch (error) {
      console.error('保存配置失败:', error)
      return { success: false, error: error.message }
    }
  })

  // 测试推送
  ipcMain.handle('test-push', async (event, type, data) => {
    const config = store.store
    if (type === 'stats') {
      const month = data?.month || null
      const stats = excelHandler.getStats(month)
      return pushHandler.sendStatsMessage(stats, config)
    } else if (type === 'celebration') {
      return pushHandler.sendCelebrationMessage(data, config)
    }
  })

  // 小程序轮询获取未同步数据
  ipcMain.handle('get-unsynced-data', async () => {
    try {
      return JSON.parse(JSON.stringify(excelHandler.getUnsyncedData()))
    } catch (e) {
      return []
    }
  })

  // 标记已同步
  ipcMain.handle('mark-synced', async (event, ids) => {
    try {
      return JSON.parse(JSON.stringify(excelHandler.markSynced(ids)))
    } catch (e) {
      return { success: false, error: e.message }
    }
  })

  // 导入业绩表
  ipcMain.handle('import-performance', async (event, filePath) => {
    try {
      return JSON.parse(JSON.stringify(excelHandler.importPerformance(filePath)))
    } catch (e) {
      return { success: false, error: e.message }
    }
  })

  // 导出业绩表
  ipcMain.handle('export-performance', async (event, outputPath, year, month) => {
    try {
      return JSON.parse(JSON.stringify(excelHandler.exportPerformance(outputPath, year, month)))
    } catch (e) {
      return { success: false, error: e.message }
    }
  })

  // 打开文件选择对话框
  ipcMain.handle('show-open-dialog', async (event, options) => {
    return dialog.showOpenDialog(mainWindow, options)
  })

  // 打开保存文件对话框
  ipcMain.handle('show-save-dialog', async (event, options) => {
    return dialog.showSaveDialog(mainWindow, options)
  })

  // 获取月份列表
  ipcMain.handle('get-month-list', async () => {
    return excelHandler.getMonthList()
  })

  // 设置当前月份
  ipcMain.handle('set-current-month', async (event, month) => {
    return excelHandler.setCurrentMonth(month)
  })

  // 读取指定月份的数据
  ipcMain.handle('get-performance-data-by-month', async (event, month) => {
    try {
      return JSON.parse(JSON.stringify(excelHandler.readPerformanceData(month)))
    } catch (e) {
      return { success: false, error: e.message }
    }
  })

  // 获取指定月份的统计
  ipcMain.handle('get-stats-by-month', async (event, month, filters) => {
    try {
      return JSON.parse(JSON.stringify(excelHandler.getStats(month, filters)))
    } catch (e) {
      return { total: 0, byTeacher: {}, count: 0 }
    }
  })

  // 添加月份
  ipcMain.handle('add-month', async (event, newSheetName, templateSheetName) => {
    try {
      return JSON.parse(JSON.stringify(excelHandler.addMonth(newSheetName, templateSheetName)))
    } catch (e) {
      return { success: false, error: e.message }
    }
  })

  // 删除月份
  ipcMain.handle('delete-month', async (event, sheetName) => {
    try {
      return JSON.parse(JSON.stringify(excelHandler.deleteMonth(sheetName)))
    } catch (e) {
      return { success: false, error: e.message }
    }
  })
}

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow()
  }
})

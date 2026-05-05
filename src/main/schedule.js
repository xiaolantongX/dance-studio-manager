const schedule = require('node-schedule')

let scheduledJobs = {}

// 初始化定时任务
function initSchedule(pushHandler) {
  // 从配置文件加载定时任务
  // 初始调用时可能还没有配置，后续通过 updateSchedule 更新
  console.log('定时任务模块已初始化')
}

// 更新/创建定时任务
function updateSchedule(config, pushHandler) {
  // 清除旧的定时任务
  Object.keys(scheduledJobs).forEach(key => {
    if (scheduledJobs[key]) {
      scheduledJobs[key].cancel()
    }
  })
  scheduledJobs = {}

  if (!config?.schedule?.enabled) {
    console.log('定时任务未启用')
    return
  }

  const times = config.schedule.times || []
  if (times.length === 0) {
    console.log('未配置推送时间')
    return
  }

  const frequency = config.schedule.frequency || 'daily'
  const weekDays = config.schedule.weekdays || []
  const monthDay = config.schedule.monthDay || 1

  // 为每个配置的时间创建定时任务
  times.forEach((time, index) => {
    const [hour, minute] = time.split(':').map(Number)

    let rule = new schedule.RecurrenceRule()
    rule.hour = hour
    rule.minute = minute
    rule.second = 0

    // 根据频率设置不同的规则
    if (frequency === 'weekly') {
      // 每周的特定几天 (node-schedule: 0=周日，1=周一，... 6=周六)
      rule.dayOfWeek = weekDays.map(d => (d === 0 ? 7 : d)) // 转换为用户格式（1-7，周一到周日）
    } else if (frequency === 'monthly') {
      rule.date = monthDay
    }
    // daily 频率不需要额外设置

    const jobKey = `${frequency}_${index}_${time}`
    const job = schedule.scheduleJob(rule, async () => {
      console.log('执行定时统计推送:', new Date())

      // 引入 excel 模块获取最新数据
      const excelHandler = require('./excel')
      const stats = excelHandler.getStats()

      await pushHandler.sendStatsMessage(stats, config)
    })

    scheduledJobs[jobKey] = job
    console.log(`已创建定时任务 [${frequency}]: ${time}`)
  })
}

// 手动触发统计推送
function triggerDailyPush(config, pushHandler) {
  const excelHandler = require('./excel')
  const stats = excelHandler.getStats()
  return pushHandler.sendStatsMessage(stats, config)
}

module.exports = {
  initSchedule,
  updateSchedule,
  triggerDailyPush
}

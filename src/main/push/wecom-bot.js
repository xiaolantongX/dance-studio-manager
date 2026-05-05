const axios = require('axios')

function getChannelConfig(config) {
  return config?.push?.wecomBot || config?.wechat || {}
}

// 发送统计消息到企业微信群
async function sendStatsMessage(stats, config) {
  try {
    const channelConfig = getChannelConfig(config)
    const webhookUrl = channelConfig.webhookUrl
    if (!webhookUrl) {
      return { success: false, error: '未配置企业微信机器人 Webhook' }
    }

    let teacherStats = stats.byTeacher || {}

    const selectedTeachers = (channelConfig.selectedTeachers && channelConfig.selectedTeachers.length > 0)
      ? channelConfig.selectedTeachers
      : (config?.push?.selectedTeachers || [])

    if (selectedTeachers.length > 0) {
      const filtered = {}
      selectedTeachers.forEach(teacher => {
        if (teacherStats[teacher] !== undefined) {
          filtered[teacher] = teacherStats[teacher]
        }
      })
      teacherStats = filtered
    }

    const sorted = Object.entries(teacherStats).sort((a, b) => b[1] - a[1])

    const monthLabel = stats.monthLabel || '全部月份'
    let content = `【${monthLabel}】\n总业绩：${stats.total.toFixed(1)}\n`
    sorted.forEach(([teacher, amount], index) => {
      let prefix = ''
      if (index === 0) prefix = '🏆'
      else if (index === 1) prefix = '🥈'
      else if (index === 2) prefix = '🥉'
      content += `${prefix}${teacher}:${amount.toFixed(1)}\n`
    })

    const payload = {
      msgtype: 'markdown',
      markdown: { content }
    }

    const response = await axios.post(webhookUrl, payload)
    if (response.data.errcode === 0) {
      return { success: true, channel: 'wecom-bot' }
    }
    return { success: false, error: response.data.errmsg, channel: 'wecom-bot' }
  } catch (error) {
    console.error('企微机器人发送失败:', error)
    return { success: false, error: error.message, channel: 'wecom-bot' }
  }
}

// 发送开卡喜报
async function sendCelebrationMessage(data, config) {
  try {
    const channelConfig = getChannelConfig(config)
    const webhookUrl = channelConfig.webhookUrl
    if (!webhookUrl) {
      return { success: false, error: '未配置企业微信机器人 Webhook' }
    }

    const template = channelConfig.celebrationTemplate || config?.push?.celebrationTemplate || defaultCelebrationTemplate

    let content = template
      .replace(/\{teacher\}/g, data.teacher || '')
      .replace(/\{cardType\}/g, data.cardType || '')
      .replace(/\{amount\}/g, data.amount || '')

    const payload = {
      msgtype: 'markdown',
      markdown: { content }
    }

    const response = await axios.post(webhookUrl, payload)
    if (response.data.errcode === 0) {
      return { success: true, channel: 'wecom-bot' }
    }
    return { success: false, error: response.data.errmsg, channel: 'wecom-bot' }
  } catch (error) {
    console.error('企微机器人喜报发送失败:', error)
    return { success: false, error: error.message, channel: 'wecom-bot' }
  }
}

const defaultCelebrationTemplate = `@所有人 [玫瑰]🌹
         传喜讯
        最新喜讯
星巢舞蹈
播报恭喜   {teacher} 老师

帮助  {cardType}一张
👏👏👏努力当下👏👏👏👏👏🌹🌹🌹🌹`

module.exports = {
  sendStatsMessage,
  sendCelebrationMessage,
  defaultCelebrationTemplate
}

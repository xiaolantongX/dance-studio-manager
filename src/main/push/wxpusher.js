const axios = require('axios')

const API_BASE = 'https://wxpusher.zjiecode.com/api'

function getChannelConfig(config) {
  return config?.push?.wxpusher || {}
}

// 发送统计消息到微信（通过 WxPusher）
async function sendStatsMessage(stats, config) {
  try {
    const channelConfig = getChannelConfig(config)
    const appToken = channelConfig.appToken
    const topicIds = channelConfig.topicIds || []

    if (!appToken) {
      return { success: false, error: '未配置 WxPusher AppToken' }
    }
    if (topicIds.length === 0) {
      return { success: false, error: '未配置 WxPusher 主题 ID' }
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
    let content = `## ${monthLabel}业绩统计\n\n`
    content += `**总业绩：${stats.total.toFixed(1)}**\n\n`
    sorted.forEach(([teacher, amount], index) => {
      let prefix = ''
      if (index === 0) prefix = '🏆 '
      else if (index === 1) prefix = '🥈 '
      else if (index === 2) prefix = '🥉 '
      content += `- ${prefix}${teacher}：${amount.toFixed(1)}\n`
    })

    const payload = {
      appToken,
      content,
      contentType: 3, // 3 = markdown
      topicIds,
      summary: `总业绩：${stats.total.toFixed(1)}`
    }

    const response = await axios.post(`${API_BASE}/send/message`, payload, {
      headers: { 'Content-Type': 'application/json' }
    })

    if (response.data.code === 1000) {
      return { success: true, channel: 'wxpusher' }
    }
    return { success: false, error: response.data.msg || '发送失败', channel: 'wxpusher' }
  } catch (error) {
    console.error('WxPusher 发送失败:', error)
    return { success: false, error: error.message, channel: 'wxpusher' }
  }
}

// 发送开卡喜报（WxPusher 版本）
async function sendCelebrationMessage(data, config) {
  try {
    const channelConfig = getChannelConfig(config)
    const appToken = channelConfig.appToken
    const topicIds = channelConfig.topicIds || []

    if (!appToken) {
      return { success: false, error: '未配置 WxPusher AppToken' }
    }
    if (topicIds.length === 0) {
      return { success: false, error: '未配置 WxPusher 主题 ID' }
    }

    const template = channelConfig.celebrationTemplate || config?.push?.celebrationTemplate ||
      `🎉 喜报 🎉\n\n恭喜 {teacher} 老师\n帮助客户办理 {cardType} 一张\n金额：{amount}\n\n👏👏👏努力当下👏👏👏`

    let content = template
      .replace(/\{teacher\}/g, data.teacher || '')
      .replace(/\{cardType\}/g, data.cardType || '')
      .replace(/\{amount\}/g, data.amount || '')

    const payload = {
      appToken,
      content,
      contentType: 3,
      topicIds,
      summary: `恭喜 ${data.teacher || ''} 老师 - ${data.cardType || ''}`
    }

    const response = await axios.post(`${API_BASE}/send/message`, payload, {
      headers: { 'Content-Type': 'application/json' }
    })

    if (response.data.code === 1000) {
      return { success: true, channel: 'wxpusher' }
    }
    return { success: false, error: response.data.msg || '发送失败', channel: 'wxpusher' }
  } catch (error) {
    console.error('WxPusher 喜报发送失败:', error)
    return { success: false, error: error.message, channel: 'wxpusher' }
  }
}

module.exports = {
  sendStatsMessage,
  sendCelebrationMessage
}

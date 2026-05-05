const crypto = require('crypto')
const axios = require('axios')

function getChannelConfig(config) {
  return config?.push?.dingtalk || {}
}

// HMAC-SHA256 加签
function generateSign(secret, timestamp) {
  const signStr = `${timestamp}\n${secret}`
  return crypto.createHmac('sha256', secret).update(signStr).digest('base64')
}

// 构建带签名的 webhook URL
function buildWebhookUrl(channelConfig) {
  const baseUrl = channelConfig.webhookUrl
  if (!baseUrl) return null

  const secret = channelConfig.secret
  if (!secret) return baseUrl

  const timestamp = Date.now()
  const sign = encodeURIComponent(generateSign(secret, timestamp))
  return `${baseUrl}&timestamp=${timestamp}&sign=${sign}`
}

// 发送统计消息到钉钉群
async function sendStatsMessage(stats, config) {
  try {
    const channelConfig = getChannelConfig(config)
    const webhookUrl = buildWebhookUrl(channelConfig)
    if (!webhookUrl) {
      return { success: false, error: '未配置钉钉机器人 Webhook' }
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
      msgtype: 'markdown',
      markdown: {
        title: '业绩统计',
        text: content
      },
      at: { isAtAll: true }
    }

    const response = await axios.post(webhookUrl, payload)
    if (response.data.errcode === 0) {
      return { success: true, channel: 'dingtalk' }
    }
    return { success: false, error: response.data.errmsg, channel: 'dingtalk' }
  } catch (error) {
    console.error('钉钉机器人发送失败:', error)
    return { success: false, error: error.message, channel: 'dingtalk' }
  }
}

// 发送开卡喜报到钉钉群
async function sendCelebrationMessage(data, config) {
  try {
    const channelConfig = getChannelConfig(config)
    const webhookUrl = buildWebhookUrl(channelConfig)
    if (!webhookUrl) {
      return { success: false, error: '未配置钉钉机器人 Webhook' }
    }

    const template = channelConfig.celebrationTemplate || config?.push?.celebrationTemplate ||
      `🎉 **喜报** 🎉\n\n恭喜 **{teacher}** 老师\n帮助客户办理 **{cardType}** 一张\n金额：**{amount}**\n\n👏👏👏努力当下👏👏👏`

    let content = template
      .replace(/\{teacher\}/g, data.teacher || '')
      .replace(/\{cardType\}/g, data.cardType || '')
      .replace(/\{amount\}/g, data.amount || '')

    const payload = {
      msgtype: 'markdown',
      markdown: {
        title: '喜报',
        text: content
      },
      at: { isAtAll: true }
    }

    const response = await axios.post(webhookUrl, payload)
    if (response.data.errcode === 0) {
      return { success: true, channel: 'dingtalk' }
    }
    return { success: false, error: response.data.errmsg, channel: 'dingtalk' }
  } catch (error) {
    console.error('钉钉机器人喜报发送失败:', error)
    return { success: false, error: error.message, channel: 'dingtalk' }
  }
}

module.exports = {
  sendStatsMessage,
  sendCelebrationMessage
}

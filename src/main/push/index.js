const wecomBot = require('./wecom-bot')
const wxpusher = require('./wxpusher')
const dingtalk = require('./dingtalk')

function isChannelEnabled(config, channel) {
  if (channel === 'wecom-bot') {
    const c = config?.push?.wecomBot || config?.wechat
    return !!(c?.webhookUrl)
  }
  if (channel === 'wxpusher') {
    const c = config?.push?.wxpusher
    return !!(c?.appToken && c?.topicIds?.length > 0)
  }
  if (channel === 'dingtalk') {
    const c = config?.push?.dingtalk
    return !!(c?.webhookUrl)
  }
  return false
}

const CHANNELS = [
  { name: 'wecom-bot', handler: wecomBot },
  { name: 'wxpusher', handler: wxpusher },
  { name: 'dingtalk', handler: dingtalk }
]

// 发送统计消息到所有启用的通道
async function sendStatsMessage(stats, config) {
  const results = []

  for (const ch of CHANNELS) {
    if (isChannelEnabled(config, ch.name)) {
      results.push(await ch.handler.sendStatsMessage(stats, config))
    }
  }

  if (results.length === 0) {
    return { success: false, error: '未配置任何推送通道' }
  }

  const anySuccess = results.some(r => r.success)
  const errors = results.filter(r => !r.success).map(r => r.error).join('; ')
  return {
    success: anySuccess,
    error: anySuccess ? null : errors,
    results,
    summary: results.map(r => `${r.channel}: ${r.success ? '成功' : r.error}`).join('; ')
  }
}

// 发送喜报到所有启用的通道
async function sendCelebrationMessage(data, config) {
  const results = []

  for (const ch of CHANNELS) {
    if (isChannelEnabled(config, ch.name)) {
      results.push(await ch.handler.sendCelebrationMessage(data, config))
    }
  }

  if (results.length === 0) {
    return { success: false, error: '未配置任何推送通道' }
  }

  const anySuccess = results.some(r => r.success)
  const errors = results.filter(r => !r.success).map(r => r.error).join('; ')
  return {
    success: anySuccess,
    error: anySuccess ? null : errors,
    results,
    summary: results.map(r => `${r.channel}: ${r.success ? '成功' : r.error}`).join('; ')
  }
}

module.exports = {
  sendStatsMessage,
  sendCelebrationMessage,
  isChannelEnabled
}

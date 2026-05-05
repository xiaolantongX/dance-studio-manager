const axios = require('axios')

// 发送统计消息到微信群
async function sendStatsMessage(stats, config) {
  try {
    const webhookUrl = config?.wechat?.webhookUrl
    if (!webhookUrl) {
      return { success: false, error: '未配置企业微信机器人 Webhook' }
    }

    // 格式化业绩数据，支持按老师筛选
    let teacherStats = stats.byTeacher || {}

    // 如果配置了只推送特定老师
    if (config.wechat?.selectedTeachers && config.wechat.selectedTeachers.length > 0) {
      const filtered = {}
      config.wechat.selectedTeachers.forEach(teacher => {
        if (teacherStats[teacher] !== undefined) {
          filtered[teacher] = teacherStats[teacher]
        }
      })
      teacherStats = filtered
    }

    // 排序（从高到低）
    const sorted = Object.entries(teacherStats)
      .sort((a, b) => b[1] - a[1])

    // 构建消息内容
    let content = `@所有人\n`
    content += `总业绩：${stats.total.toFixed(1)}\n`

    sorted.forEach(([teacher, amount], index) => {
      // 前三名加奖牌
      let prefix = ''
      if (index === 0) prefix = '🏆'
      else if (index === 1) prefix = '🥈'
      else if (index === 2) prefix = '🥉'

      content += `${teacher}${prefix}:${amount.toFixed(1)}\n`
    })

    // 企业微信markdown 格式
    const payload = {
      msgtype: 'markdown',
      markdown: {
        content: content
      }
    }

    const response = await axios.post(webhookUrl, payload)

    if (response.data.errcode === 0) {
      return { success: true }
    } else {
      return { success: false, error: response.data.errmsg }
    }
  } catch (error) {
    console.error('发送统计消息失败:', error)
    return { success: false, error: error.message }
  }
}

// 发送开卡喜报
async function sendCelebrationMessage(data, config) {
  try {
    const webhookUrl = config?.wechat?.webhookUrl
    if (!webhookUrl) {
      return { success: false, error: '未配置企业微信机器人 Webhook' }
    }

    // 使用配置的模板，如果没有则使用默认模板
    const template = config?.wechat?.celebrationTemplate || defaultCelebrationTemplate

    console.log('接收到的数据:', data)
    console.log('原始模板:', template)

    // 替换模板变量
    let content = template
      .replace(/\{teacher\}/g, data.teacher || '')
      .replace(/\{cardType\}/g, data.cardType || '')
      .replace(/\{amount\}/g, data.amount || '')

    console.log('替换后的内容:', content)

    const payload = {
      msgtype: 'markdown',
      markdown: {
        content: content
      }
    }

    const response = await axios.post(webhookUrl, payload)

    if (response.data.errcode === 0) {
      return { success: true }
    } else {
      return { success: false, error: response.data.errmsg }
    }
  } catch (error) {
    console.error('发送喜报消息失败:', error)
    return { success: false, error: error.message }
  }
}

// 默认喜报模板
const defaultCelebrationTemplate = `@所有人 [玫瑰]🌹
         传喜讯
        最新喜讯
星巢舞蹈
播报恭喜   {teacher} 老师

帮助  {cardType}一张
👏👏👏努力当下👏👏👏👏👏[🌹🌹🌹🌹`

// 发送文本消息（通用）
async function sendTextMessage(text, config) {
  try {
    const webhookUrl = config?.wechat?.webhookUrl
    if (!webhookUrl) {
      return { success: false, error: '未配置 Webhook' }
    }

    const payload = {
      msgtype: 'text',
      text: {
        content: text,
        mentioned_list: ['@all']
      }
    }

    const response = await axios.post(webhookUrl, payload)

    if (response.data.errcode === 0) {
      return { success: true }
    } else {
      return { success: false, error: response.data.errmsg }
    }
  } catch (error) {
    console.error('发送文本消息失败:', error)
    return { success: false, error: error.message }
  }
}

module.exports = {
  sendStatsMessage,
  sendCelebrationMessage,
  sendTextMessage,
  defaultCelebrationTemplate
}

<template>
  <div class="app">
    <!-- 顶部提示 -->
    <transition name="toast">
      <div v-if="toast.show" class="toast" :class="toast.type">
        {{ toast.message }}
      </div>
    </transition>

    <!-- 月份选择栏 -->
    <div class="month-selector">
      <span class="month-label">月份：</span>
      <select class="select month-select" v-model="currentMonth" @change="switchMonth">
        <option value="">全部</option>
        <option v-for="month in monthList" :key="month" :value="month">{{ month }}</option>
      </select>
      <button class="btn btn-primary" @click="openMonthManager" style="margin-left: 10px; padding: 8px 16px;">
        月份管理
      </button>
      <span v-if="excelBinding.fileName" style="margin-left: 20px; color: #909399; font-size: 13px;">
        当前文件：{{ excelBinding.fileName }}
      </span>
    </div>

    <div class="tabs">
      <div
        class="tab"
        :class="{ active: currentTab === 'dashboard' }"
        @click="currentTab = 'dashboard'"
      >
        数据概览
      </div>
      <div
        class="tab"
        :class="{ active: currentTab === 'performance' }"
        @click="currentTab = 'performance'"
      >
        业绩管理
      </div>
      <div
        class="tab"
        :class="{ active: currentTab === 'basic' }"
        @click="currentTab = 'basic'"
      >
        基础数据
      </div>
      <div
        class="tab"
        :class="{ active: currentTab === 'config' }"
        @click="currentTab = 'config'"
      >
        配置
      </div>
    </div>

    <div class="container">
      <!-- 数据概览 -->
      <div v-if="currentTab === 'dashboard'">
        <div class="card">
          <div class="card-title">业绩统计<span v-if="currentMonth" class="month-tag">{{ currentMonth }}</span><span v-else class="month-tag all">全部月份</span></div>
          <div class="stats-header">
            <div class="stats-grid">
              <div class="stat-card">
                <div class="stat-label">{{ currentMonth ? currentMonth + '总业绩' : '总业绩' }}</div>
                <div class="stat-value">¥{{ stats.total.toFixed(1) }}</div>
                <div class="stat-label">共 {{ stats.count }} 条记录</div>
              </div>
              <div class="stat-card green">
                <div class="stat-label">今日业绩</div>
                <div class="stat-value">¥{{ todayStats.total.toFixed(1) }}</div>
                <div class="stat-label">{{ todayStats.count }} 条记录</div>
              </div>
            </div>
            <button class="btn btn-primary push-btn" @click="pushStatsNow">
              推送统计数据
            </button>
          </div>

          <div class="teacher-stats" v-if="Object.keys(stats.byTeacher).length > 0">
            <div class="card-title" style="margin-top: 20px;">当月个人业绩排行</div>
            <div class="teacher-item" v-for="(amount, teacher) in sortedTeacherStats" :key="teacher">
              <span class="teacher-name teacher-link" @click="filterByTeacher(teacher)">{{ teacher }}</span>
              <span class="teacher-amount">¥{{ amount.toFixed(1) }}</span>
            </div>
          </div>

          <div class="teacher-stats" v-if="Object.keys(todayStats.byTeacher).length > 0">
            <div class="card-title" style="margin-top: 20px;">今日个人业绩</div>
            <div class="teacher-item" v-for="(amount, teacher) in sortedTodayTeacherStats" :key="teacher">
              <span class="teacher-name teacher-link" @click="filterByTeacher(teacher)">{{ teacher }}</span>
              <span class="teacher-amount">¥{{ amount.toFixed(1) }}</span>
            </div>
          </div>
        </div>
      </div>

      <!-- 业绩管理 -->
      <div v-if="currentTab === 'performance'">
        <div class="card">
          <div class="card-title">新增业绩</div>
          <div class="form-row">
            <div class="form-item" style="margin: 0;">
              <label>客户姓名</label>
              <input type="text" class="input" v-model="newPerformance.customerName" placeholder="客户姓名">
            </div>
            <div class="form-item" style="margin: 0;">
              <label>顾问</label>
              <select class="select" v-model="newPerformance.teacher">
                <option value="">选择顾问</option>
                <option v-for="t in teacherList" :key="t" :value="t">{{ t }}</option>
              </select>
            </div>
            <div class="form-item" style="margin: 0;">
              <label>卡种</label>
              <select class="select" v-model="newPerformance.cardType">
                <option value="">选择卡种</option>
                <option v-for="c in cardTypeList" :key="c" :value="c">{{ c }}</option>
              </select>
            </div>
            <div class="form-item" style="margin: 0;">
              <label>金额</label>
              <input :key="amountInputKey" type="text" class="input" v-model="newPerformance.amount" placeholder="0.00" inputmode="decimal">
            </div>
            <button class="btn btn-primary" @click="addPerformance" style="height: 42px;">添加</button>
          </div>
          <div class="form-item" style="margin-top: 10px;">
            <label>备注</label>
            <input type="text" class="input" v-model="newPerformance.remark" placeholder="可选">
          </div>
        </div>

        <div class="card">
          <div class="card-title">业绩列表</div>

          <!-- 过滤条件 -->
          <div class="filter-bar">
            <div class="filter-item">
              <label>顾问</label>
              <select class="select filter-select" v-model="filters.teacher">
                <option value="">全部顾问</option>
                <option v-for="t in teacherList" :key="t" :value="t">{{ t }}</option>
              </select>
            </div>
            <div class="filter-item">
              <label>卡种</label>
              <select class="select filter-select" v-model="filters.cardType">
                <option value="">全部卡种</option>
                <option v-for="c in cardTypeList" :key="c" :value="c">{{ c }}</option>
              </select>
            </div>
            <div class="filter-item">
              <button class="btn btn-cancel" @click="clearFilters" style="height: 42px; margin-top: auto;">
                清除过滤
              </button>
            </div>
          </div>

          <table class="table">
            <thead>
              <tr>
                <th class="sortable" @click="toggleSort('date')">
                  日期 <span class="sort-arrow">{{ sortField === 'date' ? (sortAsc ? '▲' : '▼') : '' }}</span>
                </th>
                <th>客户姓名</th>
                <th>顾问</th>
                <th>卡种</th>
                <th class="sortable" @click="toggleSort('amount')">
                  金额 <span class="sort-arrow">{{ sortField === 'amount' ? (sortAsc ? '▲' : '▼') : '' }}</span>
                </th>
                <th>备注</th>
                <th>操作</th>
              </tr>
            </thead>
            <tbody>
              <tr v-for="(item, index) in filteredPerformanceData" :key="index" :class="{ 'pending-row': item.__pending }">
                <td>{{ item['日期'] }}<span v-if="item.__pending" class="pending-tag">待同步</span></td>
                <td>{{ getCustomerName(item) }}</td>
                <td><span class="teacher-link" @click="filterByTeacher(getAdvisorName(item))">{{ getAdvisorName(item) }}</span></td>
                <td>{{ item['卡种'] }}</td>
                <td>¥{{ parseFloat(item['金额'] || 0).toFixed(1) }}</td>
                <td>{{ item['备注'] || '-' }}</td>
                <td>
                  <button class="btn btn-danger" @click="deletePerformance(item)">删除</button>
                </td>
              </tr>
              <tr v-if="performanceData.length === 0">
                <td colspan="7" style="text-align: center; color: #999; padding: 40px;">暂无数据</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      <!-- 基础数据管理 -->
      <div v-if="currentTab === 'basic'">
        <div class="card">
          <div class="card-title">顾问管理</div>
          <div style="display: flex; gap: 10px; margin-bottom: 15px;">
            <input
              type="text"
              class="input"
              v-model="newTeacher"
              placeholder="输入顾问姓名"
              @keyup.enter="addTeacher"
            >
            <button class="btn btn-primary" @click="addTeacher" style="white-space: nowrap;">添加顾问</button>
          </div>
          <div class="tag-list">
            <span class="tag" v-for="(teacher, index) in teacherList" :key="teacher">
              {{ teacher }}
              <button class="tag-close" @click="deleteTeacher(index)">×</button>
            </span>
            <span v-if="teacherList.length === 0" style="color: #999;">暂无顾问，请添加</span>
          </div>
        </div>

        <div class="card">
          <div class="card-title">卡种管理</div>
          <div style="display: flex; gap: 10px; margin-bottom: 15px;">
            <input
              type="text"
              class="input"
              v-model="newCardType"
              placeholder="输入卡种名称"
              @keyup.enter="addCardType"
            >
            <button class="btn btn-primary" @click="addCardType" style="white-space: nowrap;">添加卡种</button>
          </div>
          <div class="tag-list">
            <span class="tag" v-for="(cardType, index) in cardTypeList" :key="cardType">
              {{ cardType }}
              <button class="tag-close" @click="deleteCardType(index)">×</button>
            </span>
            <span v-if="cardTypeList.length === 0" style="color: #999;">暂无卡种，请添加</span>
          </div>
        </div>
      </div>

      <!-- 配置 -->
      <div v-if="currentTab === 'config'">
        <div class="card">
          <div class="card-title">Excel 文件绑定</div>
          <div class="form-item">
            <label>当前绑定文件</label>
            <div style="display: flex; align-items: center; gap: 10px;">
              <input
                type="text"
                class="input"
                :value="excelBinding.filePath || '未绑定'"
                readonly
                style="flex: 1;"
              >
              <button class="btn btn-primary" @click="bindExcelFile">选择文件</button>
            </div>
            <div style="margin-top: 8px; font-size: 12px; color: #909399;">
              绑定后，桌面端将直接读写该 Excel 文件，手动修改文件后会自动同步
            </div>
          </div>
        </div>

        <div class="card">
          <div class="card-title">推送通道配置</div>

          <!-- 通用喜报设置 -->
          <div class="form-item">
            <label>喜报模板（所有通道共用）</label>
            <textarea class="input" v-model="config.push.celebrationTemplate" rows="4"
              style="width: 100%; resize: vertical;"
              :placeholder="'@所有人 [玫瑰]🌹\n星巢舞蹈\n播报恭喜 {teacher} 老师\n帮助 {cardType}一张\n👏👏👏努力当下👏👏👏'"></textarea>
            <div style="margin-top: 8px; font-size: 12px; color: #909399;">
              可用变量：{teacher} = 顾问姓名，{cardType} = 卡种名称，{amount} = 金额
            </div>
          </div>
          <div class="form-item">
            <label>推送顾问选择（空=全部，所有通道共用）</label>
            <div class="checkbox-group">
              <label class="checkbox-item" v-for="teacher in allTeachers" :key="teacher">
                <input type="checkbox" :value="teacher" v-model="config.push.selectedTeachers">{{ teacher }}
              </label>
            </div>
          </div>

          <!-- 企业微信群机器人 -->
          <div class="channel-section">
            <div class="channel-header" @click="channelExpanded.wecom = !channelExpanded.wecom" style="cursor: pointer; display: flex; align-items: center; gap: 8px; margin-bottom: 12px;">
              <span style="font-size: 16px;">{{ channelExpanded.wecom ? '▼' : '▶' }}</span>
              <span style="font-weight: 600; color: #e4e6eb;">企业微信群机器人</span>
              <span v-if="config.wechat.webhookUrl" style="font-size: 12px; color: #67c23a;">已配置 ✓</span>
              <span v-else style="font-size: 12px; color: #909399;">未配置</span>
            </div>
            <div v-if="channelExpanded.wecom">
              <div class="form-item">
                <label>Webhook URL</label>
                <input type="text" class="input" v-model="config.wechat.webhookUrl"
                  placeholder="https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=xxx">
                <div style="margin-top: 8px; font-size: 12px; color: #909399;">
                  在企业微信群中添加群机器人，复制 Webhook 地址
                </div>
              </div>
            </div>
          </div>

          <!-- WxPusher -->
          <div class="channel-section" style="margin-top: 16px; padding-top: 16px; border-top: 1px solid rgba(255,255,255,0.06);">
            <div class="channel-header" @click="channelExpanded.wxpusher = !channelExpanded.wxpusher" style="cursor: pointer; display: flex; align-items: center; gap: 8px; margin-bottom: 12px;">
              <span style="font-size: 16px;">{{ channelExpanded.wxpusher ? '▼' : '▶' }}</span>
              <span style="font-weight: 600; color: #e4e6eb;">WxPusher（微信推送）</span>
              <span v-if="config.push.wxpusher.appToken" style="font-size: 12px; color: #67c23a;">已配置 ✓</span>
              <span v-else style="font-size: 12px; color: #909399;">未配置</span>
            </div>
            <div v-if="channelExpanded.wxpusher">
              <div class="form-item">
                <label>AppToken</label>
                <input type="text" class="input" v-model="config.push.wxpusher.appToken" placeholder="AT_xxxxxxxxxxxxxxxxxxxx">
                <div style="margin-top: 8px; font-size: 12px; color: #909399;">
                  在 <a href="https://wxpusher.zjiecode.com" target="_blank" style="color: #409eff;">wxpusher.zjiecode.com</a> 创建应用获取 AppToken
                </div>
              </div>
              <div class="form-item">
                <label>主题 ID（多个用逗号分隔）</label>
                <input type="text" class="input"
                  :value="config.push.wxpusher.topicIds.join(',')"
                  @input="config.push.wxpusher.topicIds = $event.target.value.split(',').map(s => parseInt(s.trim())).filter(n => !isNaN(n))"
                  placeholder="123,456">
                <div style="margin-top: 8px; font-size: 12px; color: #909399;">在 WxPusher 后台创建主题，老师扫码关注后即可收到推送</div>
              </div>
            </div>
          </div>

          <!-- 钉钉群机器人 -->
          <div class="channel-section" style="margin-top: 16px; padding-top: 16px; border-top: 1px solid rgba(255,255,255,0.06);">
            <div class="channel-header" @click="channelExpanded.dingtalk = !channelExpanded.dingtalk" style="cursor: pointer; display: flex; align-items: center; gap: 8px; margin-bottom: 12px;">
              <span style="font-size: 16px;">{{ channelExpanded.dingtalk ? '▼' : '▶' }}</span>
              <span style="font-weight: 600; color: #e4e6eb;">钉钉群机器人</span>
              <span v-if="config.push.dingtalk.webhookUrl" style="font-size: 12px; color: #67c23a;">已配置 ✓</span>
              <span v-else style="font-size: 12px; color: #909399;">未配置</span>
            </div>
            <div v-if="channelExpanded.dingtalk">
              <div class="form-item">
                <label>Webhook URL</label>
                <input type="text" class="input" v-model="config.push.dingtalk.webhookUrl"
                  placeholder="https://oapi.dingtalk.com/robot/send?access_token=xxx">
                <div style="margin-top: 8px; font-size: 12px; color: #909399;">
                  在钉钉群设置 → 智能群助手 → 添加机器人 → 自定义，复制 Webhook 地址
                </div>
              </div>
              <div class="form-item">
                <label>加签密钥（可选）</label>
                <input type="text" class="input" v-model="config.push.dingtalk.secret" placeholder="SECxxxxxxxxxx">
                <div style="margin-top: 8px; font-size: 12px; color: #909399;">
                  如果机器人的安全设置选择了"加签"，需填入对应的密钥
                </div>
              </div>
            </div>
          </div>

          <div class="form-item" style="margin-top: 16px; padding-top: 16px; border-top: 1px solid rgba(255,255,255,0.06);">
            <label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
              <input type="checkbox" v-model="config.push.pushStatsOnAdd">
              开卡后自动推送当月业绩统计
            </label>
            <div style="margin-top: 4px; font-size: 12px; color: #909399;">
              开启后，每次新增业绩将先推送喜报，再推送更新后的当月业绩排行
            </div>
          </div>
        </div>

        <div class="card">
          <div class="card-title">数据导出</div>
          <div class="import-export-section">
            <div class="ie-item">
              <h4>导出业绩表</h4>
              <p class="ie-desc">导出指定年份/月份的业绩数据</p>
              <div class="export-options">
                <select class="select" v-model="exportYear" style="width: auto;">
                  <option value="">选择年份</option>
                  <option v-for="y in yearOptions" :key="y" :value="y">{{ y }}年</option>
                </select>
                <select class="select" v-model="exportMonth" style="width: auto;">
                  <option value="">全年</option>
                  <option v-for="m in 12" :key="m" :value="m">{{ m.toString().padStart(2, '0') }}月</option>
                </select>
                <button class="btn btn-success" @click="exportPerformance">导出数据</button>
              </div>
            </div>
          </div>
        </div>

        <div class="card">
          <div class="card-title">定时任务配置</div>
          <div class="form-item">
            <label>
              <input type="checkbox" v-model="config.schedule.enabled"> 启用定时推送
            </label>
          </div>

          <div v-if="config.schedule.enabled">
            <div class="form-item">
              <label>推送频率</label>
              <select class="select" v-model="config.schedule.frequency">
                <option value="daily">每天</option>
                <option value="weekly">每周</option>
                <option value="monthly">每月</option>
              </select>
            </div>

            <div class="form-item" v-if="config.schedule.frequency === 'weekly'">
              <label>选择星期</label>
              <div class="checkbox-group">
                <label class="checkbox-item" v-for="(day, index) in weekDays" :key="index">
                  <input type="checkbox" :value="index" v-model="config.schedule.weekdays">
                  {{ day }}
                </label>
              </div>
            </div>

            <div class="form-item" v-if="config.schedule.frequency === 'monthly'">
              <label>选择日期</label>
              <select class="select" v-model="config.schedule.monthDay">
                <option v-for="day in 31" :key="day" :value="day">{{ day }}日</option>
              </select>
            </div>

            <div class="form-item">
              <label>推送时间</label>
              <div style="display: flex; gap: 10px; align-items: center; margin-bottom: 10px;">
                <input type="time" class="input" v-model="scheduleTime" style="width: auto;">
                <button class="btn btn-primary" @click="addScheduleTime" style="padding: 8px 16px;">添加时间</button>
              </div>
              <div class="time-tags">
                <span class="time-tag" v-for="(time, index) in config.schedule.times" :key="index">
                  {{ time }}
                  <button class="time-tag-close" @click="removeScheduleTime(index)">×</button>
                </span>
                <span v-if="!config.schedule.times || config.schedule.times.length === 0" style="color: #999;">暂无推送时间，请添加</span>
              </div>
            </div>
          </div>
        </div>

        <div class="card">
          <button class="btn btn-primary" @click="saveConfig">保存配置</button>
        </div>
      </div>
    </div>
  </div>

  <!-- 确认对话框 -->
  <transition name="fade">
    <div v-if="confirmDialog.show" class="modal-overlay" style="z-index: 9999;">
      <div class="modal">
        <div class="modal-message">{{ confirmDialog.message }}</div>
        <div class="modal-buttons">
          <button class="btn btn-cancel" @click="confirmNo()">取消</button>
          <button class="btn btn-primary" @click="confirmYes()">确定</button>
        </div>
      </div>
    </div>
  </transition>

  <!-- 月份管理对话框 -->
  <transition name="fade">
    <div v-if="monthManager.show" class="modal-overlay">
      <div class="modal month-manager-modal">
        <div class="card-title" style="margin-top: 0;">月份管理</div>
        <div class="month-add-row" style="display: flex; gap: 10px; margin-bottom: 16px;">
          <input
            type="text"
            class="input"
            v-model="monthManager.newMonthName"
            placeholder="输入新月份名称，如：2026年6月"
            @keyup.enter="addMonth"
            style="flex: 1;"
          >
          <button class="btn btn-primary" @click="addMonth" style="white-space: nowrap;">添加月份</button>
        </div>
        <div class="month-list" style="max-height: 300px; overflow-y: auto;">
          <div class="month-list-item" v-for="month in monthList" :key="month" style="display: flex; justify-content: space-between; align-items: center; padding: 10px 12px; border-bottom: 1px solid #ebeef5;">
            <span>{{ month }}</span>
            <button class="btn btn-danger btn-sm" @click="deleteMonthByName(month)" :disabled="monthList.length <= 1" style="padding: 4px 12px; font-size: 12px;">删除</button>
          </div>
          <div v-if="monthList.length === 0" style="text-align: center; color: #999; padding: 20px;">暂无月份</div>
        </div>
        <div class="modal-buttons" style="margin-top: 16px;">
          <button class="btn btn-cancel" @click="monthManager.show = false">关闭</button>
        </div>
      </div>
    </div>
  </transition>
</template>

<script>
const { ipcRenderer } = require('electron')

export default {
  name: 'App',
  data() {
    return {
      currentTab: 'dashboard',
      performanceData: [],
      stats: { total: 0, byTeacher: {}, count: 0 },
      todayStats: { total: 0, byTeacher: {}, count: 0 },
      teacherList: [],
      cardTypeList: [],
      allTeachers: [],
      newTeacher: '',
      newCardType: '',
      amountInputKey: 0,
      tempScheduleTime: '',
      filters: {
        teacher: '',
        cardType: ''
      },
      sortField: '',
      sortAsc: false,
      exportYear: '',
      exportMonth: '',
      yearOptions: [],
      currentMonth: '',
      monthList: [],
      pendingCount: 0,
      excelBinding: {
        filePath: '',
        fileName: '',
        exists: false
      },
      toast: {
        show: false,
        message: '',
        type: 'info'
      },
      confirmDialog: {
        show: false,
        message: '',
        callback: null
      },
      monthManager: {
        show: false,
        newMonthName: ''
      },
      channelExpanded: {
        wecom: true,
        wxpusher: false,
        dingtalk: false
      },
      pendingPerformance: null,
      itemsToAdd: {
        teacher: false,
        cardType: false
      },
      newPerformance: {
        customerName: '',
        teacher: '',
        cardType: '',
        amount: '',
        remark: ''
      },
      config: {
        wechat: {
          webhookUrl: '',
          celebrationTemplate: '',
          selectedTeachers: []
        },
        push: {
          pushStatsOnAdd: false,
          celebrationTemplate: '',
          selectedTeachers: [],
          wxpusher: {
            enabled: false,
            appToken: '',
            topicIds: [],
            selectedTeachers: []
          },
          dingtalk: {
            enabled: false,
            webhookUrl: '',
            secret: '',
            selectedTeachers: []
          }
        },
        schedule: {
          enabled: false,
          frequency: 'daily',
          dailyTime: '21:00',
          intervalEnabled: false,
          intervalMinutes: 60,
          weekdays: [],
          monthDay: 1,
          times: []
        }
      }
    }
  },
  computed: {
    weekDays() {
      return ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
    },
    scheduleTime: {
      get() {
        return this.tempScheduleTime || ''
      },
      set(val) {
        this.tempScheduleTime = val
      }
    },
    sortedTeacherStats() {
      return Object.entries(this.stats.byTeacher)
        .sort((a, b) => b[1] - a[1])
        .reduce((acc, [k, v]) => ({ ...acc, [k]: v }), {})
    },
    sortedTodayTeacherStats() {
      return Object.entries(this.todayStats.byTeacher)
        .sort((a, b) => b[1] - a[1])
        .reduce((acc, [k, v]) => ({ ...acc, [k]: v }), {})
    },
    filteredPerformanceData() {
      let data = this.performanceData

      if (this.filters.teacher) {
        data = data.filter(item => this.getAdvisorName(item) === this.filters.teacher)
      }
      if (this.filters.cardType) {
        data = data.filter(item => item['卡种'] === this.filters.cardType)
      }

      if (this.sortField) {
        data = [...data].sort((a, b) => {
          if (this.sortField === 'amount') {
            const diff = (parseFloat(b['金额']) || 0) - (parseFloat(a['金额']) || 0)
            return this.sortAsc ? -diff : diff
          }
          // sort by date - use the timestamp for reliable comparison
          const dateA = a['日期'] || ''
          const dateB = b['日期'] || ''
          const cmp = String(dateA).localeCompare(String(dateB))
          return this.sortAsc ? -cmp : cmp
        })
      }

      return data
    }
  },
  watch: {
    currentTab(newVal) {
      if (newVal === 'performance') {
        this.refreshData()
      }
    }
  },
  async mounted() {
    this.refreshData()
    this.loadConfig()
    this.loadExcelBindingInfo()
    // 初始化学份选项
    const currentYear = new Date().getFullYear()
    this.yearOptions = [currentYear, currentYear - 1, currentYear - 2, currentYear + 1]
    // 获取月份列表并自动定位到当前月份
    await this.loadMonthList()
    await this.autoSelectCurrentMonth()
    // 监听 Excel 文件变化
    ipcRenderer.on('excel-file-changed', this.handleExcelFileChanged)
  },
  beforeUnmount() {
    ipcRenderer.removeListener('excel-file-changed', this.handleExcelFileChanged)
  },
  methods: {
    getCustomerName(item) {
      return item['姓名'] || item['客户姓名'] || ''
    },
    getAdvisorName(item) {
      return item['顾问'] || item['老师'] || ''
    },
    async loadExcelBindingInfo() {
      const binding = await ipcRenderer.invoke('get-excel-binding-info')
      if (binding) {
        this.excelBinding = binding
      }
    },
    async bindExcelFile() {
      const result = await ipcRenderer.invoke('show-open-dialog', {
        title: '选择 Excel 文件',
        filters: [
          { name: 'Excel 文件', extensions: ['xlsx', 'xls'] }
        ],
        properties: ['openFile']
      })

      if (result.canceled || result.filePaths.length === 0) {
        return
      }

      const filePath = result.filePaths[0]
      const bindResult = await ipcRenderer.invoke('bind-excel-file', filePath)
      if (bindResult.success) {
        this.showToast('绑定成功', 'success')
        this.excelBinding = bindResult.binding
        this.loadMonthList()
        this.refreshData()
      } else {
        this.showToast('绑定失败：' + bindResult.error, 'error')
      }
    },
    handleExcelFileChanged(event, payload) {
      console.log('Excel 文件变化:', payload)
      this.loadExcelBindingInfo()
      this.loadMonthList()
      this.refreshData()
    },
    async loadMonthList() {
      const months = await ipcRenderer.invoke('get-month-list')
      this.monthList = months
    },
    async autoSelectCurrentMonth() {
      const now = new Date()
      const currentMonthName = `${now.getFullYear()}年${now.getMonth() + 1}月`

      if (!this.monthList.includes(currentMonthName)) {
        // 自动创建当月 sheet，以最新月份的表头为模板
        const result = await ipcRenderer.invoke('add-month', currentMonthName, this.monthList[this.monthList.length - 1] || null)
        if (result.success) {
          this.monthList.push(currentMonthName)
        }
      }

      this.currentMonth = currentMonthName
      await ipcRenderer.invoke('set-current-month', this.currentMonth)
      this.loadMonthData()
    },
    async switchMonth() {
      await ipcRenderer.invoke('set-current-month', this.currentMonth || null)
      this.loadMonthData()
    },
    async loadMonthData() {
      this.performanceData = []
      this.stats = { total: 0, byTeacher: {}, count: 0 }
      await this.refreshData()
    },
    openMonthManager() {
      this.monthManager.show = true
      this.monthManager.newMonthName = ''
    },
    async addMonth() {
      const name = this.monthManager.newMonthName.trim()
      if (!name) {
        this.showToast('请输入月份名称', 'error')
        return
      }
      const result = await ipcRenderer.invoke('add-month', name, this.currentMonth || null)
      if (result.success) {
        this.showToast(`月份"${name}"添加成功`, 'success')
        this.monthManager.newMonthName = ''
        await this.loadMonthList()
      } else {
        this.showToast('添加失败：' + result.error, 'error')
      }
    },
    async deleteMonthByName(month) {
      if (this.monthList.length <= 1) {
        this.showToast('至少保留一个月份', 'error')
        return
      }
      this.showConfirm(`确定删除月份"${month}"及其所有数据吗？此操作不可恢复。`, async () => {
        const result = await ipcRenderer.invoke('delete-month', month)
        if (result.success) {
          this.showToast(`月份"${month}"已删除`, 'success')
          if (this.currentMonth === month) {
            this.currentMonth = ''
          }
          await this.loadMonthList()
          this.refreshData()
        } else {
          this.showToast('删除失败：' + result.error, 'error')
        }
      })
    },
    showToast(message, type = 'success') {
      this.toast = { show: true, message, type }
      const duration = type === 'info' ? 5000 : 2000
      setTimeout(() => {
        this.toast.show = false
      }, duration)
    },
    showConfirm(message, callback) {
      this.confirmDialog = { show: true, message, callback }
    },
    formatChineseDate(date = new Date()) {
      return `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}号`
    },
    confirmYes() {
      if (this.confirmDialog.callback) {
        this.confirmDialog.callback()
      }
      this.confirmDialog.show = false
    },
    confirmNo() {
      this.confirmDialog.show = false
    },
    pushStatsNow() {
      this.testPush('stats', { month: this.currentMonth || null })
    },
    clearFilters() {
      this.filters.teacher = ''
      this.filters.cardType = ''
      this.sortField = ''
    },
    toggleSort(field) {
      if (this.sortField === field) {
        this.sortAsc = !this.sortAsc
      } else {
        this.sortField = field
        this.sortAsc = false
      }
    },
    filterByTeacher(teacher) {
      this.filters.teacher = teacher
      this.currentTab = 'performance'
    },
    deletePerformance(record) {
      this.showConfirm('确定删除这条业绩吗？', async () => {
        const result = await ipcRenderer.invoke('delete-performance', record)
        if (result.success) {
          if (result.pending) {
            this.showToast(result.message || '删除已暂存，关闭 Excel 后将自动同步', 'info')
          } else {
            this.showToast('删除成功', 'success')
          }
          this.refreshData()
        } else {
          this.showToast('删除失败：' + result.error, 'error')
        }
      })
    },
    async refreshData() {
      console.log('refreshData called, currentMonth:', this.currentMonth)
      // 读取业绩数据（支持按月份）
      const month = this.currentMonth || null
      const result = await ipcRenderer.invoke('get-performance-data-by-month', month)
      console.log('get-performance-data-by-month result:', result)
      if (result.success) {
        this.performanceData = result.data || []
      } else {
        this.performanceData = []
        if (result.error && result.error.includes('不存在')) {
          this.showToast('Excel 文件不存在，请重新绑定', 'error')
        }
      }

      // 获取统计（支持按月份）
      // 注意：this.filters 是 Vue 响应式对象，IPC 无法克隆，需转为普通对象
      const statsResult = await ipcRenderer.invoke('get-stats-by-month', month, { teacher: this.filters.teacher, cardType: this.filters.cardType })
      console.log('get-stats-by-month result:', statsResult)
      this.stats = statsResult

      // 从配置中获取老师和卡种列表（如果没有配置则从数据中提取）
      const config = await ipcRenderer.invoke('get-config')
      if (config && config.basic) {
        if (config.basic.teacherList && config.basic.teacherList.length > 0) {
          this.teacherList = config.basic.teacherList
        } else {
          this.teacherList = [...new Set((result.data || []).map(item => this.getAdvisorName(item)).filter(Boolean))]
        }
        if (config.basic.cardTypeList && config.basic.cardTypeList.length > 0) {
          this.cardTypeList = config.basic.cardTypeList
        } else {
          this.cardTypeList = [...new Set((result.data || []).map(item => item['卡种']).filter(Boolean))]
        }
        this.allTeachers = this.teacherList
      } else {
        this.teacherList = [...new Set((result.data || []).map(item => this.getAdvisorName(item)).filter(Boolean))]
        this.cardTypeList = [...new Set((result.data || []).map(item => item['卡种']).filter(Boolean))]
        this.allTeachers = this.teacherList
      }

      // 今日统计使用后端计算结果（基于 isSameDate 精确匹配）
      this.todayStats = {
        total: statsResult.todayTotal || 0,
        byTeacher: statsResult.todayByTeacher || {},
        count: statsResult.todayCount || 0
      }

      // 提示待处理更改
      const prevPending = this.pendingCount
      this.pendingCount = result.pendingCount || 0
      if (this.pendingCount > 0) {
        this.showToast(`有 ${this.pendingCount} 条更改等待同步（Excel 文件被占用），关闭 Excel 后自动同步`, 'info')
      } else if (prevPending > 0 && result.flushed) {
        this.showToast('待处理更改已成功同步到 Excel', 'success')
      }
    },
    async addPerformance() {
      if (!this.newPerformance.teacher || !this.newPerformance.cardType || !this.newPerformance.amount) {
        this.showToast('请填写完整信息', 'error')
        return
      }

      // 检查老师和卡种是否在数据管理中
      const teacherExists = this.teacherList.includes(this.newPerformance.teacher)
      const cardTypeExists = this.cardTypeList.includes(this.newPerformance.cardType)

      // 如果不存在，需要添加
      if (!teacherExists || !cardTypeExists) {
        const itemsToAdd = []
        if (!teacherExists) itemsToAdd.push(`顾问"${this.newPerformance.teacher}"`)
        if (!cardTypeExists) itemsToAdd.push(`卡种"${this.newPerformance.cardType}"`)

        this.pendingPerformance = {
          customerName: this.newPerformance.customerName || '',
          teacher: this.newPerformance.teacher,
          cardType: this.newPerformance.cardType,
          amount: parseFloat(this.newPerformance.amount) || 0,
          date: this.formatChineseDate()
        }
        this.itemsToAdd = {
          teacher: !teacherExists,
          cardType: !cardTypeExists
        }

        this.showConfirm(`检测到 ${itemsToAdd.join('、')} 不在数据管理中，是否要添加到数据管理？`, () => {
          this.confirmAddAndSave()
        })
        return
      }

      // 直接添加业绩
      await this.savePerformance()
    },
    async confirmAddAndSave() {
      // 添加到老师列表
      if (this.itemsToAdd.teacher && this.pendingPerformance.teacher) {
        if (!this.teacherList.includes(this.pendingPerformance.teacher)) {
          this.teacherList.push(this.pendingPerformance.teacher)
        }
      }
      // 添加到卡种列表
      if (this.itemsToAdd.cardType && this.pendingPerformance.cardType) {
        if (!this.cardTypeList.includes(this.pendingPerformance.cardType)) {
          this.cardTypeList.push(this.pendingPerformance.cardType)
        }
      }
      // 保存配置
      this.saveBasicConfig()
      // 保存业绩
      await this.savePerformance()
      this.pendingPerformance = null
      this.itemsToAdd = { teacher: false, cardType: false }
    },
    async savePerformance() {
      const performanceData = this.pendingPerformance || {
        customerName: this.newPerformance.customerName || '',
        teacher: this.newPerformance.teacher,
        cardType: this.newPerformance.cardType,
        amount: parseFloat(this.newPerformance.amount) || 0,
        date: this.formatChineseDate()
      }

      const result = await ipcRenderer.invoke('add-performance', performanceData, this.currentMonth || null)

      if (result.success) {
        // 待处理模式：Excel 被锁定，暂存成功但未写入文件
        if (result.pending) {
          this.showToast(result.message || '业绩已暂存，关闭 Excel 后将自动写入', 'info')
        } else {
          // 添加成功后自动推送喜报
          const config = await ipcRenderer.invoke('get-config')
          const hasPushChannel = config && (
            (config.wechat && config.wechat.webhookUrl) ||
            (config.push && config.push.wxpusher && config.push.wxpusher.appToken) ||
            (config.push && config.push.dingtalk && config.push.dingtalk.webhookUrl)
          )
          if (hasPushChannel) {
            await ipcRenderer.invoke('test-push', 'celebration', {
              teacher: performanceData.teacher,
              cardType: performanceData.cardType,
              amount: performanceData.amount
            })
            // 开卡后自动推送当月业绩统计
            if (config.push && config.push.pushStatsOnAdd) {
              await ipcRenderer.invoke('test-push', 'stats', { month: this.currentMonth || null })
            }
          }
          this.showToast('添加成功', 'success')
        }
        // 逐个属性重置，保持响应式
        this.newPerformance.customerName = ''
        this.newPerformance.teacher = ''
        this.newPerformance.cardType = ''
        this.newPerformance.amount = ''
        this.newPerformance.remark = ''
        // 增加 key 值强制重新创建输入框
        this.amountInputKey++
        this.refreshData()
      } else {
        this.showToast('添加失败：' + result.error, 'error')
      }
    },
    async loadConfig() {
      const config = await ipcRenderer.invoke('get-config')
      if (config) {
        // 深度合并 push 配置以保留默认值
        if (config.push) {
          if (config.push.pushStatsOnAdd !== undefined) {
            this.config.push.pushStatsOnAdd = config.push.pushStatsOnAdd
          }
          if (config.push.celebrationTemplate !== undefined) {
            this.config.push.celebrationTemplate = config.push.celebrationTemplate
          }
          if (config.push.selectedTeachers !== undefined) {
            this.config.push.selectedTeachers = config.push.selectedTeachers
          }
          if (config.push.wxpusher) {
            Object.assign(this.config.push.wxpusher, config.push.wxpusher)
          }
          if (config.push.dingtalk) {
            Object.assign(this.config.push.dingtalk, config.push.dingtalk)
          }
          delete config.push
        }
        this.config = { ...this.config, ...config }
        // 从配置中加载老师和卡种列表
        if (config.basic && config.basic.teacherList) {
          this.teacherList = config.basic.teacherList
        }
        if (config.basic && config.basic.cardTypeList) {
          this.cardTypeList = config.basic.cardTypeList
        }
        // 设置默认模板（共享配置）
        if (!this.config.push.celebrationTemplate) {
          this.config.push.celebrationTemplate = `@所有人 [玫瑰]🌹
         传喜讯
        最新喜讯
星巢舞蹈
播报恭喜   {teacher} 老师

帮助  {cardType}一张
👏👏👏努力当下👏👏👏👏👏🌹🌹🌹🌹`
        }
      }
      this.allTeachers = this.teacherList
    },
    addTeacher() {
      if (!this.newTeacher.trim()) {
        this.showToast('请输入顾问姓名', 'error')
        return
      }
      if (this.teacherList.includes(this.newTeacher.trim())) {
        this.showToast('该顾问已存在', 'error')
        return
      }
      this.teacherList.push(this.newTeacher.trim())
      this.newTeacher = ''
      this.saveBasicConfig()
    },
    deleteTeacher(index) {
      this.showConfirm(`确定删除顾问"${this.teacherList[index]}"吗？`, () => {
        this.teacherList.splice(index, 1)
        this.saveBasicConfig()
      })
    },
    addCardType() {
      if (!this.newCardType.trim()) {
        this.showToast('请输入卡种名称', 'error')
        return
      }
      if (this.cardTypeList.includes(this.newCardType.trim())) {
        this.showToast('该卡种已存在', 'error')
        return
      }
      this.cardTypeList.push(this.newCardType.trim())
      this.newCardType = ''
      this.saveBasicConfig()
    },
    deleteCardType(index) {
      this.showConfirm(`确定删除卡种"${this.cardTypeList[index]}"吗？`, () => {
        this.cardTypeList.splice(index, 1)
        this.saveBasicConfig()
      })
    },
    saveBasicConfig() {
      this.config.basic = {
        teacherList: this.teacherList,
        cardTypeList: this.cardTypeList
      }
      const configToSave = JSON.parse(JSON.stringify(this.config))
      ipcRenderer.invoke('save-config', configToSave)
    },
    addScheduleTime() {
      if (!this.tempScheduleTime) {
        this.showToast('请选择时间', 'error')
        return
      }
      if (!this.config.schedule.times) {
        this.config.schedule.times = []
      }
      if (this.config.schedule.times.includes(this.tempScheduleTime)) {
        this.showToast('该时间已存在', 'error')
        return
      }
      this.config.schedule.times.push(this.tempScheduleTime)
      this.config.schedule.times.sort()
      this.tempScheduleTime = ''
    },
    removeScheduleTime(index) {
      this.config.schedule.times.splice(index, 1)
    },
    async saveConfig() {
      // 使用 JSON 序列化创建一个纯对象，避免 Vue 响应式对象无法克隆
      const configToSave = JSON.parse(JSON.stringify(this.config))
      const result = await ipcRenderer.invoke('save-config', configToSave)
      if (result.success) {
        this.showToast('保存成功', 'success')

        // 如果启用了定时推送，显示下次推送时间
        if (this.config.schedule.enabled && this.config.schedule.times && this.config.schedule.times.length > 0) {
          const nextPushTime = this.getNextPushTime()
          if (nextPushTime) {
            setTimeout(() => {
              this.showToast(`下次推送时间：${nextPushTime}`, 'info')
            }, 2000)
          }
        }
      } else {
        this.showToast('保存失败', 'error')
      }
    },
    getNextPushTime() {
      const now = new Date()
      const times = this.config.schedule.times || []
      if (times.length === 0) return null

      // 找到下一个推送时间
      for (let time of times) {
        const [hours, minutes] = time.split(':').map(Number)
        const pushTime = new Date()
        pushTime.setHours(hours, minutes, 0, 0)
        if (pushTime > now) {
          return pushTime.toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit' })
        }
      }

      // 如果今天的推送时间已过，返回明天的第一个时间
      return times[0]
    },
    async testPush(type, extra) {
      const result = await ipcRenderer.invoke('test-push', type, extra)
      if (result.success) {
        this.showToast('推送成功', 'success')
      } else {
        this.showToast('推送失败：' + (result.error || '未知错误'), 'error')
      }
    },
    async importPerformance() {
      const result = await ipcRenderer.invoke('show-open-dialog', {
        title: '选择业绩表文件',
        filters: [
          { name: 'Excel 文件', extensions: ['xlsx', 'xls'] }
        ],
        properties: ['openFile']
      })

      if (result.canceled || result.filePaths.length === 0) {
        return
      }

      const filePath = result.filePaths[0]
      this.showToast('正在导入...', 'info')

      const importResult = await ipcRenderer.invoke('import-performance', filePath)
      if (importResult.success) {
        this.showToast(`导入成功，共 ${importResult.count} 条记录，${importResult.months ? importResult.months.length : ''} 个月份`, 'success')
        this.loadMonthList()
        this.refreshData()
      } else {
        this.showToast('导入失败：' + importResult.error, 'error')
      }
    },
    async exportPerformance() {
      if (!this.exportYear) {
        this.showToast('请选择年份', 'error')
        return
      }

      const result = await ipcRenderer.invoke('show-save-dialog', {
        title: '保存业绩表',
        filters: [
          { name: 'Excel 文件', extensions: ['xlsx'] }
        ],
        defaultPath: `${this.exportYear}年${this.exportMonth ? this.exportMonth + '月' : '全年'}业绩表.xlsx`
      })

      if (result.canceled || !result.filePath) {
        return
      }

      const exportResult = await ipcRenderer.invoke('export-performance', result.filePath, this.exportYear, this.exportMonth)
      if (exportResult.success) {
        this.showToast(`导出成功，共 ${exportResult.count} 条记录`, 'success')
      } else {
        this.showToast('导出失败：' + exportResult.error, 'error')
      }
    }
  }
}
</script>

<style scoped>
.app {
  min-height: 100vh;
}
</style>

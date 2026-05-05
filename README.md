# 舞蹈室业绩管理系统

Electron + Vue3 桌面端应用，用于舞蹈工作室的业绩录入、统计与自动推送。

## 功能

- **数据概览** — 当月总业绩 / 今日业绩，按老师排行，支持月份切换
- **业绩管理** — 增删业绩记录，按顾问/卡种筛选，金额/日期排序
- **月份管理** — 自动创建当月 Sheet，切换历史月份
- **多通道推送** — 企业微信群机器人 / 钉钉群机器人 / WxPusher(微信)
- **定时推送** — 可配置每日/每周/每月自动推送业绩统计
- **喜报推送** — 开卡后自动推送喜报，模板可自定义
- **待处理队列** — Excel 被占用时操作暂存，关闭后自动同步

## 技术栈

| 层 | 技术 |
|---|---|
| 框架 | Electron 28 |
| 前端 | Vue 3 + Vite 5 |
| 数据 | xlsx (直接读写 Excel) |
| 推送 | axios → 企微/钉钉 Webhook / WxPusher API |
| 存储 | electron-store (配置持久化) |
| 打包 | electron-builder + vite-plugin-singlefile |

## 快速开始

### 环境要求

- Node.js 18+
- Windows 10/11

### 安装运行

```bash
# 安装依赖
npm install

# 开发模式
npm run electron:dev

# 打包为 exe
npm run electron:build
```

## 推送通道配置

### 企业微信群机器人

1. 企业微信群 → 设置 → 群机器人 → 添加自定义机器人
2. 复制 Webhook URL：`https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=xxx`
3. 填入系统配置页 → 保存

### 钉钉群机器人

1. 钉钉群 → 设置 → 智能群助手 → 添加机器人 → 自定义
2. 复制 Webhook URL：`https://oapi.dingtalk.com/robot/send?access_token=xxx`
3. 如设置了加签安全方式，填入对应密钥
4. 填入系统配置页 → 保存

### WxPusher（微信推送）

1. 访问 [wxpusher.zjiecode.com](https://wxpusher.zjiecode.com) 创建应用
2. 获取 AppToken，创建主题获取 TopicId
3. 让老师微信扫码关注主题
4. 填入系统配置页 → 保存

## 目录结构

```
desktop/
├── src/
│   ├── main/                  # Electron 主进程
│   │   ├── main.js            # 窗口管理、IPC 处理
│   │   ├── excel.js           # Excel 读写与业务逻辑
│   │   ├── schedule.js        # 定时任务调度
│   │   └── push/              # 推送通道模块
│   │       ├── index.js       # 统一路由
│   │       ├── wecom-bot.js   # 企业微信
│   │       ├── dingtalk.js    # 钉钉
│   │       └── wxpusher.js    # WxPusher
│   └── renderer/              # Vue3 前端
│       ├── main.js            # Vue 入口
│       ├── App.vue            # 单文件组件
│       └── style.css          # 全局样式
├── index.html                 # HTML 入口
├── vite.config.js             # Vite 构建配置
└── package.json
```

## 更多功能

- **喜报模板**：支持 `{teacher}` `{cardType}` `{amount}` 变量
- **老师筛选**：可选择只推送特定老师的业绩
- **开卡后自动推当月统计**：可在配置中开启

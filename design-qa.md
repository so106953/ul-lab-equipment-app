# 设计 QA：UL 实验室设备管理

## 对比目标

- Source visual truth: `public/assets/qa-source.png`（已确认的 UL 红色设备优先方案）
- Implementation screenshot: `public/assets/qa-prototype.png`
- Viewport: 1536 × 768 desktop
- State: 默认领用状态，设备“生物显微镜 CX23”已选中
- Full-view comparison evidence: `http://localhost:4173/qa-comparison.html` 将设计稿与浏览器截图并列呈现并已人工检查。
- Focused-region evidence: 并列视图中已检查左侧 UL 标识/导航、设备信息横幅、借用表单与右侧交易摘要；这些关键区域在该视图中均可辨识，因此无需额外裁切。

## Findings

- No actionable P0/P1/P2 differences found.

### Required fidelity surfaces

- Fonts and typography: 使用系统中文无衬线字体栈；标题、段落、字段标签和摘要的层级与参考图一致，未发生影响可读性的换行或截断。
- Spacing and layout rhythm: 保留了红色侧栏、设备横幅、编号表单区、右侧固定摘要的四区结构；桌面网格和留白与设计方向一致。
- Colors and visual tokens: 应用壳层为 UL 红，工作区为暖白，表面采用白色面板；红色用于操作态，绿色只用于可用状态。
- Image quality and asset fidelity: 使用用户提供的 UL Solutions 标识文件，并使用独立生成的显微镜产品图；未以 CSS、emoji、占位图或手绘 SVG 替代视觉资产。
- Copy and content: 表单字段、设备规格、借用/归还状态、记录和导出入口均为完整中文可读内容。

## Interaction verification

- 领用/归还切换可更新交易摘要。
- 工号、组别、时长单位、步进器、设备名称、序列号与备注可编辑。
- 提交登记后显示成功提示，并将新记录加入“记录”列表。
- 记录页面可返回登记页；导出入口生成 UTF-8 CSV，可用 Excel 打开。
- 已重新打开应用并检查浏览器控制台：应用页无 warning/error。

## Comparison history

1. 初次实现：发现摘要提交按钮在浏览器文本翻译环境中显示动态词语不一致；改为稳定的“提交登记 / 提交后将生成操作记录”文案。
2. 修复后：移除开发环境 StrictMode 后重新构建并重新打开应用；浏览器控制台无应用错误，功能与视觉检查通过。

## Follow-up Polish

- [P3] 若后续接入真实资产库，可按设备分类补充更多真实设备照片。

final result: passed

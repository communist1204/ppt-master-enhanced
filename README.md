# PPT Master 增强版

完整、独立的演示文稿制作 Skill，用于生成、美化、重构和精确编辑可编辑 PPTX，支持科研图示、学术汇报、品牌提案与成品质量审核。

当前技能版本：`6.9.0-enhanced.3-local`，基于 [Hugo He / PPT Master](https://github.com/hugohe3/ppt-master) `6.9.0`。这是个人维护的增强版本，并非上游或 Anthropic 官方版本。

## 主要改进

- 保留原版路由、脚本、模板、图标、原生对象与导出能力。
- 选择性融合 Anthropic 官方 PPTX Skill 的原生编辑与成品审核方法；支持工具为独立实现，没有复制其技能源码。
- 按页面目标选择合适的视觉表达，结合逐页渲染检查与修正；不把局部特写、多视角或固定小图数量设为通用规则。
- 增加指定对象的精确文本修改、文件完整性检查、局部编辑对比及矢量 PDF 裁切等辅助工具。

## 仓库内容

```text
skills/ppt-master/                         完整、可直接导入的技能源码
  SKILL.md                               原始技能入口
  scripts/ references/ workflows/        工具与工作流程
  templates/                             模板、图标及素材
  LICENSE SPONSORS.md SPONSORS_CN.md      原许可与归属
  requirements.txt                       运行依赖声明
dist/ppt-master-6.9.0-enhanced.3-local.zip  完整原始分发包
LICENSE                                  上游 MIT 许可证副本
```

完整 ZIP 包含 **13,104 个技能文件**，保留全部源码、模板和资源，未为上传而精简；与 `skills/ppt-master/` 对应。原包 SHA256：

```text
b09cf331175ff9e6f24a759174b3ad15ccf576f3c21d47256d095db4e316dd0e
```

使用时，将 `skills/ppt-master/` 作为 Skill 来源导入支持此类技能的代理，并从其 `SKILL.md` 开始。完整包也可独立展开使用，无需云端插件或访问原维护者电脑。

## 环境与许可

源码不包含 Python 运行时或已安装依赖。按所选任务检查 `requirements.txt` 和相关工具说明；字体、外部生图或语音服务等取决于实际环境。Windows PowerPoint 原生渲染与视频功能需要 Windows 和 Office；其他环境使用可用的替代渲染器，并明确实际使用方式。

保留上游 MIT 许可证、版权与赞助归属；模板和素材的独立许可说明随包保留。可选 PyMuPDF 依赖具有其自身许可，未随本仓库分发其运行时。完整原包保留既有 Python 缓存及其编译路径元信息，运行不应依赖这些缓存。

此前测试覆盖部分科研与品牌案例，不能保证所有未来结果都更美观；文件结构检查也不能替代实际渲染和人工视觉审核。

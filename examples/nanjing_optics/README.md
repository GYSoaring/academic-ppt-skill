# 南京大学光学超晶格样例

这个 16 页样例用 `defense` 档重建“光学超晶格与准相位匹配”答辩稿：正文约 24 pt、标题 32–40 pt、官方书法字标与校徽、真实论文图，并保留可编辑 PPTX 路线。

- 主题：`../../assets/themes/nanjing_purple/theme.json`，标识来源见 `../../assets/themes/nanjing_purple/SOURCE.md`。
- 内容：`deck.json`，16 页，`meta.genre: theory`。
- 论文图：`assets/ppln_photon_pair.png`；页 `09A` 带 `image_origin`、`source`、`image_role` 和 `image_boundary`，并设置 `image_width_ratio: 0.75` 让主图成为视觉锚点。
- 讲稿由 `build_speech.py` 生成后经过去 AI 味审计；A 级翻案腔已改写，仍应结合正式答辩口径人工复核。

复现：

```bash
python scripts/render_deck.py --deck examples/nanjing_optics/deck.json --theme assets/themes/nanjing_purple/theme.json --out work/nanjing_optics --profile defense
python scripts/export_editable_pptx.py --deck examples/nanjing_optics/deck.json --theme assets/themes/nanjing_purple/theme.json --out work/nanjing_optics --pptx-name nanjing_optics_defense.pptx --profile defense
python scripts/validate_deck.py --deck examples/nanjing_optics/deck.json --theme assets/themes/nanjing_purple/theme.json --render-dir work/nanjing_optics/rendered --out work/nanjing_optics --profile defense
python scripts/audit_pptx.py --pptx work/nanjing_optics/nanjing_optics_defense.pptx --out work/nanjing_optics
```

院校标识仅用于主题能力展示；使用时应遵循南京大学视觉标识规范，本项目与学校无隶属或背书关系。

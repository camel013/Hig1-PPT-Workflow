# PPT 运行前检查

这只做一次静态检查，不生成页面，不读取 `.venv` 或 `vendor` 源码。

```powershell
pwsh -NoProfile -File .\_ppt_workflow\preflight.ps1 -DeckName meeting_4p -Pages 4
```

它会在 `_ppt_runs/<deck>_<timestamp>/` 下建立目录并写出 `manifest.json`。后续阶段只读 manifest，不再重复查 skill、字体、Python、Edge 和参考成品路径。

审查策略固定为：

- 联系表只做快速筛查。
- 每页必须看一次原图。
- 同一页最多看一次；只有发现异常才额外看一次。

预算上限 25 次模型调用；A 阶段最多 2 次，省下的额度优先给 B/C 的设计迭代。

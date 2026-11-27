# NCC-Receipt 仓库协作与发布规范

1. **每次发布 GitHub Release 必须同时上传双包**：
   - 轻量更新包 `NCC_receipt_<版本>_update.zip`（不含内置 Chromium 浏览器，供 `latest.json` 自动更新使用）。
   - 完整离线包 `NCC_receipt_<版本>_full.zip`（含双版本 EXE、内置 Chromium、空白 33 列宏模板、空白对照表方案与说明文档）。
   - 请通过 `prepare_update_release.py --update-archive ... --full-archive ...` 与 `publish_receipt_release.py <目录> --publish` 执行校验和发布。
2. **隐私与数据保护**：
   - 仓库与公开 Release 完整包中的 `收款结算单底稿导入模板.xlsm`、`对照表方案.xlsx` 必须为 0 行业务数据的空白模板。
   - 严禁提交或打包 `.release-secrets/`、`*.dpapi`、`收款结算单填写顺序设置.json`、`收款结算单时间设置.json`。
3. **业务规则不变量**：
   - GUI 填写顺序设置仅允许表头与表头、明细与明细各自内部互换顺序，严禁表头与明细之间跨区互换。
   - `付款银行账号` 固定在所有数据（含全部明细）填妥后、点击保存/保存提交前最后一个填入。
   - 自动更新不得覆盖用户已有的底稿模板、对照表方案、填写顺序设置、时间设置及内置浏览器。

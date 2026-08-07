# 工作台接入说明

## 生产环境配置

```env
NODE_ENV=production
FRAME_ANCESTORS="'self' https://<workbench-domain>"
API_READ_TOKEN=<只读访问令牌>
API_WRITE_TOKEN=<写入访问令牌>
MAX_REQUEST_BYTES=12582912
MAX_UPLOAD_BYTES=10485760
```

- 销售人员只获得只读令牌；采购与销售支持获得写入令牌。
- 令牌必须由工作台网关注入或换取，不能写入前端代码、Git 仓库或 URL。
- 工作台采用同域反向代理时，浏览器不直接暴露后端令牌；网关负责验证登录态、数据范围与令牌注入。
- OIDC 建立后，必须以 OIDC scope 替代当前过渡性的 API 令牌。

## 验收

1. `GET /api/health` 返回 200。
2. 生产环境缺少 `FRAME_ANCESTORS` 或使用通配符时服务拒绝启动。
3. 未授权的 API、附件读取和写操作返回 401。
4. 只读令牌不能执行导入、上传或数据覆盖；写入令牌可以。
5. 上传/导入超过限制时被拒绝；COS 写入失败时 API 不返回成功。

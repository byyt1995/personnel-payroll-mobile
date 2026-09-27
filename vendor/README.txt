Microsoft 登录组件部署说明

此目录不从 CDN 加载任何脚本。

当前已从官方 npm 包 @azure/msal-browser 5.7.0 随包放入：

  mobile/vendor/msal-browser.min.js

SHA-256：9c0eecaed0eceaf10d99452b64220d0b07e6202959b5bad9bbd38363606929fe

浏览器中应提供：

  window.msal.PublicClientApplication

升级版本时只能用官方 npm 包覆盖，并同步更新本文件、MSAL-LICENSE.txt、
上述 SHA-256 和 service worker 缓存版本。若 bundle 缺失或损坏，手机页会
停止 OneDrive 登录，不会退回到自制 OAuth 或第三方 CDN。

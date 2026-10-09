# 矢印 VecStamp 宣传网站

一个纯静态的单页网站：`index.html` 内联了全部样式和脚本，不依赖任何 CDN、Google Fonts 或第三方统计，放到任何静态服务器上都能直接访问，国内服务器也不受外网资源影响。

```
website/
├── index.html          页面（样式和脚本已内联）
├── config.js           站点配置：版本号、下载地址、仓库、邮箱、备案号
├── favicon.ico
├── robots.txt
├── assets/             Logo、功能区示意图、Linux 截图、赞赏二维码、分享图
└── download/           可选：放发布包，供国内用户直接下载
```

页面包含：首屏的「截图 vs 矢量」放大对比演示、使用步骤、1.2.1 更新内容、七种格式说明、可交互的背景预览、功能区示意图与 Linux 截图、安装说明、常见问题和赞赏二维码。自动适配手机和深色模式，右上角也可以手动切换深浅色。

## 1. 修改配置

部署前一般只需要改 `config.js`：

| 字段 | 作用 |
| --- | --- |
| `version` | 页面上显示的版本号（`python build/make_website.py` 会自动同步） |
| `download` | 「下载」按钮地址。默认指向 GitHub Releases；国内访问 GitHub 较慢时，把发布包放进 `download/`，改成 `'download/VecStamp-v1.2.1.zip'` |
| `repo` | 源码仓库地址，「查看源码」「提交 Issue」「更新日志」「MIT 协议」链接都由它生成 |
| `email` | 反馈邮箱 |
| `icp` | ICP 备案号，例如 `'京ICP备12345678号-1'`。服务器在中国大陆时必须在页脚展示，留空则不显示 |
| `police` | 公安联网备案号（可选），例如 `'京公网安备11010502000000号'` |

`index.html` 里的链接都写有默认值，即使 `config.js` 没有加载，页面也能正常使用。

如果希望微信、QQ 等分享卡片显示预览图，把 `index.html` 中 `og:image` 的 `assets/og.png` 改成带域名的完整地址，例如 `https://vecstamp.example.com/assets/og.png`。

## 2. 本地预览

```bash
cd website
python3 -m http.server 8080
# 浏览器打开 http://localhost:8080
```

直接双击 `index.html` 也能打开。

## 3. 部署到云服务器

### Nginx

把整个 `website` 文件夹上传到服务器，例如 `/var/www/vecstamp`：

```bash
scp -r website/* user@your-server:/var/www/vecstamp/
```

新建 `/etc/nginx/conf.d/vecstamp.conf`：

```nginx
server {
    listen 80;
    server_name vecstamp.example.com;          # 换成你的域名

    root  /var/www/vecstamp;
    index index.html;
    charset utf-8;

    location / {
        try_files $uri $uri/ =404;
    }

    # 图片长期缓存；页面和配置每次都检查更新
    location ~* \.(png|webp|svg|ico)$ {
        expires 30d;
        add_header Cache-Control "public";
    }
    location ~* \.(html|js)$ {
        add_header Cache-Control "no-cache";
    }

    # 发布包以附件方式下载
    location /download/ {
        add_header Content-Disposition "attachment";
    }

    gzip on;
    gzip_types text/css application/javascript image/svg+xml;

    add_header X-Content-Type-Options nosniff;
    add_header Referrer-Policy strict-origin-when-cross-origin;
}
```

```bash
sudo nginx -t && sudo systemctl reload nginx
```

开启 HTTPS（Let's Encrypt 免费证书）：

```bash
sudo apt install certbot python3-certbot-nginx   # CentOS / Alibaba Cloud Linux: sudo dnf install certbot python3-certbot-nginx
sudo certbot --nginx -d vecstamp.example.com
```

也可以使用云厂商控制台签发的免费证书，下载 Nginx 格式后在 `server` 中配置 `listen 443 ssl;`、`ssl_certificate` 和 `ssl_certificate_key`。

### 宝塔面板

1. 「网站 → 添加站点」，填写域名，PHP 版本选「纯静态」。
2. 在「文件」中进入站点根目录，删除默认的 `index.html`，上传 `website` 文件夹里的全部内容。
3. 站点设置 →「SSL」申请 Let's Encrypt 证书，并开启「强制 HTTPS」。

### Caddy

```
vecstamp.example.com {
    root * /var/www/vecstamp
    file_server
    encode gzip
}
```

Caddy 会自动申请和续期 HTTPS 证书。

### 对象存储静态网站

阿里云 OSS、腾讯云 COS、七牛云等都支持「静态网站托管」：上传 `website` 文件夹的内容，把默认首页设为 `index.html`，再绑定自定义域名即可。

## 4. 备案提醒

使用中国大陆服务器并绑定域名时，需要先完成 ICP 备案，并在 `config.js` 的 `icp` 中填写备案号，页脚会自动显示并链接到工信部备案查询网站。使用香港或海外服务器不需要 ICP 备案。

## 5. 更新网站

发布新版本后，在项目根目录运行：

```bash
python build/make_promo.py      # 重新生成配图（需要 Noto Sans CJK 字体）
python build/make_website.py    # 复制配图、二维码到 website/assets，同步版本号
```

然后把 `website` 文件夹重新上传即可。页面文字（更新内容、常见问题等）直接编辑 `index.html`。

## 6. 隐私

页面不加载任何第三方脚本、字体或统计代码，不使用 Cookie；只有「切换深浅色」的选择保存在访问者自己浏览器的 localStorage 中。如果需要访问统计，建议使用服务器日志或自建统计工具。

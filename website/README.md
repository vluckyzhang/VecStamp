# 矢印 VecStamp 宣传网站

一个纯静态的单页网站，提供**简体中文（默认）、繁體中文、English、日本語**四个版本：每个页面都内联了全部样式和脚本，不依赖任何 CDN、Google Fonts 或第三方统计，放到任何静态服务器上都能直接访问，国内服务器也不受外网资源影响。

```
website/
├── index.html          简体中文页面（四种语言的源文件，样式和脚本已内联）
├── zh-TW/index.html    繁體中文页面   ┐
├── en/index.html       English 页面   ├ 由 build/make_website.py 根据 index.html 和 i18n.json 生成，不要手改
├── ja/index.html       日本語页面     ┘
├── i18n.json           页面文字的翻译表（以简体中文原文为键）
├── config.js           站点配置：版本号、网站地址、下载地址、仓库、邮箱、备案号
├── sitemap.xml         站点地图（四种语言，含 hreflang），自动生成
├── robots.txt          自动生成，指向 sitemap.xml
├── favicon.ico
├── assets/             Logo、功能区示意图、Linux 截图、赞赏二维码、分享图
└── download/           可选：放发布包，供国内用户直接下载
```

页面包含：首屏的「截图 vs 矢量」放大对比演示、使用步骤、最新版本更新内容、七种格式说明、可交互的背景预览、功能区示意图与 Linux 截图、安装说明、常见问题和赞赏二维码。自动适配手机和深色模式，右上角可以切换语言和深浅色。

## SEO（搜索引擎优化）

`build/make_website.py` 为每个语言页面写入：

- `<html lang>`、对应语言的 `<title>`、`description`、`keywords`；
- `canonical` 和全部语言的 `hreflang`（含 `x-default`），让搜索引擎把四个页面识别为同一内容的不同语言版本；
- Open Graph / Twitter 分享卡片（带域名的 `og:image`、`og:url`、`og:locale`）；
- 结构化数据（JSON-LD）：`SoftwareApplication`（名称、版本、平台、免费、下载地址、功能列表）和 `FAQPage`（来自页面上的常见问题）；
- `sitemap.xml`（四种语言及其 hreflang 对应关系）和 `robots.txt` 中的 `Sitemap:` 行。

上线后建议把 `sitemap.xml` 提交到 [百度搜索资源平台](https://ziyuan.baidu.com/)、[Google Search Console](https://search.google.com/search-console) 和 [Bing Webmaster Tools](https://www.bing.com/webmasters)。

## 1. 修改配置

部署前一般只需要改 `config.js`：

| 字段 | 作用 |
| --- | --- |
| `version` | 页面上显示的版本号（`python build/make_website.py` 会自动同步） |
| `siteUrl` | 网站的公开地址，以 `/` 结尾，例如 `'https://vecstamp.example.com/'`。canonical、hreflang、分享图、结构化数据和 sitemap 都用它生成；修改后重新运行 `python build/make_website.py --pages` |
| `download` | 「下载」按钮地址。默认指向 GitHub Releases；国内访问 GitHub 较慢时，把发布包放进 `download/`，改成 `'download/VecStamp-v1.2.2.zip'`（相对网站根目录，四种语言页面都能用） |
| `repo` | 源码仓库地址，「查看源码」「提交 Issue」「更新日志」「MIT 协议」链接都由它生成 |
| `email` | 反馈邮箱 |
| `icp` | ICP 备案号，例如 `'京ICP备12345678号-1'`。服务器在中国大陆时必须在页脚展示，留空则不显示 |
| `police` | 公安联网备案号（可选），例如 `'京公网安备11010502000000号'` |

`index.html` 里的链接都写有默认值，即使 `config.js` 没有加载，页面也能正常使用。

分享卡片的预览图 `og:image` 会按 `siteUrl` 自动写成带域名的完整地址。

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
python build/make_website.py    # 复制配图、二维码到 website/assets，同步版本号，生成四种语言页面和 sitemap
```

然后把 `website` 文件夹重新上传即可。

修改页面文字：直接编辑 `index.html`（简体中文），再在 `i18n.json` 中为新的或改动过的文字加上翻译（键是简体中文原文，值是 `zh-CN` / `zh-TW` / `en` / `ja` 四种写法）。繁体可以先用 `python build/i18n_tool.py fill-tw` 自动转换再人工检查。运行 `python build/make_website.py --pages` 时，任何缺少翻译的文字都会被列出并停止生成，不会产生半翻译的页面。

## 6. 隐私

页面不加载任何第三方脚本、字体或统计代码，不使用 Cookie；只有「切换深浅色」的选择保存在访问者自己浏览器的 localStorage 中。语言切换是普通链接，不做自动跳转。如果需要访问统计，建议使用服务器日志或自建统计工具。

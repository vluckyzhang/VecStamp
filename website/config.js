/*
 * 矢印 VecStamp 网站配置 —— 部署时通常只需要修改这个文件。
 * 页面中的链接都有默认值，即使这个文件没有加载也能正常使用。
 */
window.SITE = {
  // 当前版本号（build/make_website.py 会自动同步）
  version: '1.2.1',

  // 「下载」按钮的地址。国内访问 GitHub 较慢时，可以把发布包放到 download/ 文件夹，
  // 改成例如 'download/VecStamp-v1.2.1.zip'
  download: 'https://github.com/vluckyzhang/VecStamp/releases/latest',

  // 源码仓库（「查看源码」「提交 Issue」「更新日志」都由它生成）
  repo: 'https://github.com/vluckyzhang/VecStamp',

  // 反馈邮箱
  email: 'vluckyzhang@gmail.com',

  // ICP 备案号（服务器在中国大陆时必须在页脚展示），例如 '京ICP备12345678号-1'；留空则不显示
  icp: '',

  // 公安联网备案号（可选），例如 '京公网安备11010502000000号'；留空则不显示
  police: ''
};

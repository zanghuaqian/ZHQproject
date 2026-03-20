# 将「订单管理」页面发布到 GitHub Pages

本仓库已在 `docs/` 目录下放置静态页面（`index.html` + `.nojekyll`），与根目录 `procurement-mall-order-admin.html` 内容一致，便于 GitHub Pages 部署。

## 稳定访问链接（启用 Pages 后）

将下面地址中的 **`你的用户名`**、**`仓库名`** 换成你的 GitHub 信息：

| 说明 | URL |
|------|-----|
| **推荐（站点根路径）** | `https://你的用户名.github.io/仓库名/` |
| 直接打开同名文件 | `https://你的用户名.github.io/仓库名/procurement-mall-order-admin.html`（若根目录也有该文件且已推送） |

当前远程仓库示例：`origin` 指向 `zanghuaqian/excel-transformer-skill` 时，启用 Pages 后一般为：

`https://zanghuaqian.github.io/excel-transformer-skill/`

（具体以 GitHub 仓库 **Settings → Pages** 中显示的地址为准。）

## 一次性：推送到 GitHub

在项目根目录执行：

```bash
git add docs/
git commit -m "docs: 采购商城订单管理 GitHub Pages"
git push origin main
```

## 在 GitHub 上开启 Pages

1. 打开仓库：**GitHub → 该仓库 → Settings**
2. 左侧 **Pages**
3. **Build and deployment → Source**：选择 **Deploy from a branch**
4. **Branch**：选 `main`，文件夹选 **`/docs`**
5. 保存后等待 1～3 分钟，页面会显示 **Your site is live at …** 即为稳定链接

## 以后更新页面内容

1. 编辑仓库根目录：`procurement-mall-order-admin.html`
2. 覆盖复制到：`docs/index.html`
3. `git add` 上述文件后 `commit` 并 `push`

## 说明

- 页面依赖 **Vue / Tailwind / 字体** 的公共 CDN，无需额外构建或后端。
- 若只想单独建一个仓库放此演示，可新建空仓库，把 `docs/` 内文件拷到仓库根目录并将 `index.html` 放在根目录，Pages 选 **root** 即可。

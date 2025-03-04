# Office365 SharePoint PowerPoint 下载器

一个用于自动捕获 SharePoint 站点中 PowerPoint 演示文稿并转换为 PDF 的 Python 脚本。本项目不需要SharePoint的编辑权限，无需登录Microsoft Office365.

## 快速上手

-  [快速启动手册](easy-start-guide.md).
  
## 功能特点

- 自动导航到 SharePoint PowerPoint 演示文稿
- 以高质量截图方式捕获每一页幻灯片
- 将所有幻灯片合并为单个 PDF 文件
- 处理动画和过渡效果
- 支持全屏演示模式
- 智能检测最后一页

## 环境要求

- Python 3.7 或更高版本
- Chrome 浏览器
- 所需 Python 包：
  ```
  selenium
  Pillow
  ```

## 安装步骤

1. 克隆此仓库或下载脚本
2. 安装所需包：
   ```bash
   pip install selenium Pillow
   ```
   本项目使用版本  
   ```bash
   pip install -r requirements.txt
   ```
3. 确保已安装 Chrome 浏览器

## 使用方法

1. 运行脚本：
   ```bash
   python powerpoint_capture-zh.py
   ```

2. 根据提示输入：
   - SharePoint PowerPoint 的 URL
   - 输出文件夹名称（可选，默认为 'slides'）

3. 脚本将：
   - 在 Chrome 中打开演示文稿
   - 进入演示模式
   - 捕获每一页幻灯片
   - 将所有幻灯片保存为 PDF

## 输出内容

- 单独的幻灯片截图保存在指定文件夹中（默认：'slides'）
- 在同一文件夹中创建合并后的 PDF 文件
- 如果发生错误，会保存调试信息

## 错误处理

脚本包含健壮的错误处理机制：
- 发生错误时保存页面源码
- 重试失败的操作
- 提供详细的日志记录
- 处理网络问题和页面加载延迟

## 已知限制

- 可能无法捕获某些复杂的动画效果
- 需要网络连接
- 屏幕分辨率建议至少为 1920x1080 以获得最佳效果

## 故障排除

如果遇到问题：

1. 确保网络连接稳定
2. 检查输出文件夹中的错误日志和页面源码
3. 验证是否有正确的 PowerPoint 访问权限

## 许可证

本项目采用 MIT 许可证 - 详见 LICENSE 文件 

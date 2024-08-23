# Word助理-社区版
系统要求：Windows 7 SP1+，Office 2013+  
下载地址：[https://github.com/sinhe/word-assistant/releases/download/v1.0.0.0/setup.exe](https://github.com/sinhe/word-assistant/releases/download/v1.0.0.0/setup.exe)  
公共密钥：B9GDEnCFJ0PQB50sorpOBMnDqAsg2UcJ7f5qDURYIIU=  
功能更新：点击验证按钮即可刷新功能列表。  
数据安全：服务端仅验证密钥并提供程序运行所必须的文件，用户端输入的数据保存在本地电脑中，程序处理仅在本地Word插件环境中运行（有特殊说明的除外）。  
技术实现：使用VSTO开发Word插件并打包成exe程序安装到本机，UI采用WebBrowser加载HTML，再通过window.external调用C#代码，使用微软Office组件Microsoft.Office.Interop提供的功能处理本机Office业务逻辑。  
# ToDoList
1. ☑️批量查找替换 by fear798 （功能：批量查找和替换，保存多个条件，可添加批注，可加载Excel数据）
![1111841](https://github.com/user-attachments/assets/3f81b02c-a6af-40de-bbbd-ffd9cc35e002)
2. ☑️批量提取数据 by gix （功能：从表格中批量提取数据并导出到Excel中）
![2112243](https://github.com/user-attachments/assets/4cbc90db-ac28-46a2-b917-abc5482dc626)
3. ☑️远程智能纠错 by 流星（功能：纠正文字错误、乱码、英文和数字混淆、标点符号的全角半角错误，删除多余的换行符）此功能将发送客户端数据到服务端运行并返回结果，介意者请勿使用。

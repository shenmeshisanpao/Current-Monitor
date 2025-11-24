# 电流实时监控系统使用教程

本教程由 Z.C. Zhang 最后一次修改于2025年11月24日。

---

## 使用教程

### 1.将电流表通电，然后连接到电脑的USB口。

### 2. 在主界面"Channel Serial Port"后分别填入连接电流表的两个USB串口名，点击"Test Serial Port"按钮进行连接测试。如果只有一个电流表，可在菜单栏"Run"-"Single Channel Mode (CH1 Only)"切换到单表模式，此时仅Channel 1可用。

对于Linux系统，串口名通常是"/dev/ttyUSB\*"，其中\*是数字。

**获取USB串口名的方法如下：**

在Terminal里输入命令ls /dev/ttyUSB\*，查看已连接的USB串口。如果已连接的USB太多，难以确定哪个是电流表，还可以通过如下方法：

1.在连接电流表前，打开Terminal，输入如下命令：

```text
dmesg -w
```

2\. 连接电流表

3\. 观察输出信息，通常会显示类似：

```text
[  123.456] usb 1-1: new full-speed USB device using xhci_hcd and address 2
[  123.789] usb 1-1: New USB device found, idVendor=xxxx, idProduct=xxxx
[  124.012] usb 1-1: USB serial device converter now attached to ttyUSB0
```

这样就能看到USB设备被分配到了 `/dev/ttyUSB0` 串口。

​

### 3. 在"Save File Name"处设置记录文件的位置和文件名，默认保存在“程序文件夹/Run_0000.csv“。如果记录文件已存在，可以在"File Mode"处选择追加或覆盖，默认为追加模式。

### 4. 准备就绪，点击"Start Monitoring"按钮（Ctrl+R）开始监控。程序会自动积分电荷量、自动保存到文件。

### 5. 点击"Stop Monitoring"（Ctrl+T）按钮停止监控。如果文件名是以四位数字为结尾的话，停止监控后，会自动 +1。

### 6. 若要继续，重复第3-5步。

 

## 额外功能

1. 在电流监控运行期间，用户可以点击"Create Snapshot"按钮（快捷键Ctrl+S）保存当前记录文件的副本到记录文件同文件夹下。

2. 在菜单栏"File"-"Open Data Folder"可以打开记录文件所在的文件夹。

3. 在菜单栏"Run"-"Set Update Interval"可以设置读取电流的间隔，默认为100 ms。

4. 在菜单栏"Run"-"Pulse Reminder"可以开启/关闭打脉冲提醒，默认开启。

5. 默认为双表模式，可在菜单栏"Run"-"Single Channel Mode (CH1 Only)"切换到单表模式。

6. 鼠标悬停在曲线上可以查看该点信息。

---

​

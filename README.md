## 简介

本代码主要是根据excel的公司名称列表批量查询公司基本信息

对应的接口ID：818

## 说明

代码会根据example.xlsx的表头筛选出相应的信息存入到output.xlsx文件，所以务必保持两份文件的首行(表头)相同。

error.xlsx为无法在天眼查匹配到信息的公司列表。

comp.xlsx为待查询的公司列表，从第二行开始。

## tips：

token需要到 [天眼查官网](https://open.tianyancha.com/open/818) 申请

申请后填到代码中
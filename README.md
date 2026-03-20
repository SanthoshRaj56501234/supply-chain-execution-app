 #  Supply Chain Execution App

An interactive R Shiny dashboard that helps supply chain planners 
confirm customer orders, detect bulk orders, and manage inventory 
allocation in real time.

Business Problem
Supply chain teams struggle to quickly decide which customer orders 
can be fully confirmed vs. partially fulfilled when inventory is limited.

 What This App Does
Bulk Order Detection — flags unusual orders based on historical averages
Order Confirmation — confirms full or partial fulfillment per customer
Inventory Allocation — allocates available stock using predefined rules
Material Availability Check — validates raw material supply before committing

Technical Stack
- R, Shiny, bslib, tidyverse, DT, visNetwork, openxlsx

 Files

| `exexution_app__1_.R` | Main Shiny app code |
| `data.xlsx` | Input data (orders, BOM, supply) |

![App Demo](https://github.com/SanthoshRaj56501234/supply-chain-execution-app/blob/00e74ef47691de27cd6517a13ffe0f0854e0fc82/ScreenRecording2026-03-20at4.27.25PM-ezgif.com-video-to-gif-converter%20(1).gif)

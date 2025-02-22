# OpenAlgo Excel-DNA Add-in

This is an **Excel-DNA add-in** that integrates with **OpenAlgo API**, allowing users to retrieve financial data directly in **Microsoft Excel** using simple functions.

## 📌 Features
- **Retrieve Funds**: `=Funds()`
- **View Order Book**: `=OrderBook()`
- **Set API Configuration**: `=SetOpenAlgoConfig(api_key, version, host_url)`

## 🚀 Installation
1. **Download & Build**:
   - Clone this repository.
   - Open the project in **Visual Studio**.
   - Build the solution (`Ctrl + Shift + B`).

2. **Load the Add-in in Excel**:
   - Open **Excel**.
   - Go to **Options → Add-ins**.
   - Browse and load the generated `.xll` file from `bin\Debug` or `bin\Release`.

---

## 🔹 **Usage**

### **1️⃣ Set API Key (One Time)**
To configure the **API Key**, **version**, and **host URL**, use:
```excel
=SetOpenAlgoConfig("YOUR_API_KEY")

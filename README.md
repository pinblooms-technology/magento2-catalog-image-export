# Magento 2 - Catalog Image Export Module

## Overview
**PinBlooms_CatalogImageExport** This Magento 2 module adds a custom admin action that allows you to export selected product images directly from the product catalog grid..

## Features
- Adds "Export Products" action in Catalog → Products.
- Downloads an Excel file (.xlsx) containing product details with image references.
- No configuration required.
- Simple installation and use.

## Requirements
- PHP Spreadsheet Library: PhpSpreadsheet

## Installation
### 1. Download and Extract
Clone or download the module into your Magento 2 installation:
```sh
cd <magento_root>/app/code/PinBlooms/CatalogImageExport
```

### 2. Enable the Module
Run the following commands to enable and set up the module:
```sh
php bin/magento module:enable PinBlooms_CatalogImageExport
php bin/magento setup:upgrade
php bin/magento setup:di:compile
php bin/magento cache:flush
```
### 3. Usage
After a successful installation, you can find this module under:
- Go to Admin Panel → Catalog → Products.
- Select one or more products using the checkbox.
- From the "Actions" dropdown, choose "Export Products".
- An .xlsx file will automatically download in your browser containing exported data with image information.

  ![image (26)](https://github.com/user-attachments/assets/b7743370-f6bb-4b09-8393-fff716c66a85)




  ![image (27)](https://github.com/user-attachments/assets/cba75065-a932-4b52-a67b-de402130f5cb)


### 4. Support
For issues or feature requests, please create a GitHub issue or contact us at https://pinblooms.com/contact-us/.

### 5. License
This module is released under the MIT License.

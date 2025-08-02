#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF訂單轉Excel工具 - 英文格式批次處理版本
專門處理英文格式的PDF採購訂單
"""

import os
import sys
import argparse
import pandas as pd
import PyPDF2
from datetime import datetime, timedelta
import re
from pathlib import Path

def parse_english_order_page(text):
    """解析英文格式的訂單頁面文字"""
    orders = []
    
    # 分割成行
    lines = text.split('\n')
    
    # 尋找PO號碼
    po_match = re.search(r'PURCHASE ORDER PO#\s*(\d+)', text)
    if not po_match:
        return orders
    
    po_number = po_match.group(1)
    
    # 尋找出貨日期 - 改進的邏輯
    ship_date = ''
    # 尋找格式如 "08/01/25Country of" 的日期
    ship_date_match = re.search(r'(\d{2}/\d{2}/\d{2})Country of', text)
    if ship_date_match:
        ship_date = ship_date_match.group(1)
    
    # 尋找客戶名稱
    customer_match = re.search(r'SHIP TO\s*\n([^\n]+)', text)
    customer = 'Delta'  # 直接設為 Delta
    
    # 改進目的地提取 - 只尋找 "FOB"
    destination = 'N/A'  # 預設為 N/A
    if 'FOB' in text:
        destination = 'FOB'
    
    # 改進數量提取 - 從實際數量行提取
    quantity = 0
    # 尋找格式如 "080213072674555140 50.00 1,000" 的行
    for line in lines:
        qty_match = re.search(r'\d{12}\d+\s+\d+\.\d+\s+(\d{1,3}(?:,\d{3})*)', line)
        if qty_match:
            try:
                qty_str = qty_match.group(1).replace(',', '')
                quantity = int(qty_str)
                break
            except:
                continue
    
    # 如果沒找到，嘗試從PO TOTAL提取
    if quantity == 0:
        po_total_match = re.search(r'(\d{1,3}(?:,\d{3})*\.\d{2})\s*PO TOTAL', text)
        if po_total_match:
            try:
                quantity_str = po_total_match.group(1).replace(',', '')
                quantity = int(float(quantity_str))
            except:
                quantity = 0
    
    # 改進UPC號碼提取 - 從實際UPC行提取
    upc_number = ''
    # 尋找格式如 "080213072674555140" 的行
    upc_match = re.search(r'(\d{12})\d+', text)
    if upc_match:
        upc_number = upc_match.group(1)
    
    # 改進型號提取 - 只接受W開頭或6位數字
    model_number = ''
    # 尋找W開頭的型號
    w_model_match = re.search(r'(W\d+)', text)
    if w_model_match:
        model_number = w_model_match.group(1)
    else:
        # 尋找6位數字型號
        num_model_match = re.search(r'\b(\d{6})\b', text)
        if num_model_match:
            model_number = num_model_match.group(1)
    
    # 改進顏色代號提取 - 從SKU行提取
    color_code = ''
    # 尋找SKU行中的4位數字顏色代碼
    sku_color_matches = re.findall(r'(\d{4})\s+([A-Za-z\s]+?)\s+\d{2}/\d{2}/\d{2}', text)
    if sku_color_matches:
        for match in sku_color_matches:
            color_code_candidate = match[0]
            color_name_candidate = match[1].strip()
            # 檢查是否包含特定顏色名稱
            if any(color in color_name_candidate for color in ['WALNUT ESPRESSO', 'EBONY', 'BIANCA WHITE', 'GREY']):
                color_code = color_code_candidate
                break
    
    # 改進顏色中文提取 - 從SKU行提取
    color_name = ''
    if sku_color_matches:
        for match in sku_color_matches:
            color_code_candidate = match[0]
            color_name_candidate = match[1].strip()
            # 檢查是否包含特定顏色名稱
            if any(color in color_name_candidate for color in ['WALNUT ESPRESSO', 'EBONY', 'BIANCA WHITE', 'GREY']):
                color_name = color_name_candidate
                break
    
    # 改進LOT號碼提取 - 只接受NFVXXXXX格式，排除Reference #中的
    lot_number = ''
    # 尋找NFV開頭後跟數字的格式，但排除Reference #中的
    all_nfv_matches = re.findall(r'NFV(\d+)', text)
    for match in all_nfv_matches:
        # 檢查這個NFV數字是否在Reference #附近
        ref_context = re.search(r'Reference\s*#.*?NFV' + match, text, re.DOTALL)
        if not ref_context:
            lot_number = f"NFV{match}"
            break
    
    # 建立訂單資料
    order = {
        '客戶': customer,
        '下單日': datetime.now().strftime('%Y/%m/%d'),
        '客戶訂單出貨日': ship_date,
        'VF出貨日': '',  # 需要計算
        'PO': po_number,
        'LOT': lot_number,
        'UPC#': upc_number,
        '型號': model_number,
        '顏色代號': color_code,
        '數量': quantity,
        '目的地': destination,
        '顏色中文': color_name,
        '備註': '',
        '來源檔案': '',
        '頁面': 0
    }
    
    # 計算VF出貨日
    if ship_date:
        try:
            # 嘗試多種日期格式
            date_formats = ['%m/%d/%y', '%m/%d/%Y', '%Y/%m/%d']
            ship_dt = None
            
            for fmt in date_formats:
                try:
                    ship_dt = datetime.strptime(ship_date, fmt)
                    break
                except:
                    continue
            
            if ship_dt:
                vf_ship_dt = ship_dt - timedelta(days=3)
                order['VF出貨日'] = vf_ship_dt.strftime('%Y/%m/%d')
        except:
            order['VF出貨日'] = ''
    
    # 檢查是否有足夠的資料
    if order.get('PO') and order.get('數量'):
        orders.append(order)
    
    return orders

def process_single_pdf(pdf_path):
    """處理單個PDF檔案"""
    print(f"📄 正在處理: {os.path.basename(pdf_path)}")
    
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            all_orders = []
            
            for page_num, page in enumerate(pdf_reader.pages, 1):
                text = page.extract_text()
                
                if text:
                    orders = parse_english_order_page(text)
                    
                    if orders:
                        # 添加來源檔案和頁面資訊
                        for order in orders:
                            order['來源檔案'] = os.path.basename(pdf_path)
                            order['頁面'] = page_num
                        
                        all_orders.extend(orders)
                        print(f"  ✅ 第 {page_num} 頁: {len(orders)} 筆訂單")
                    else:
                        print(f"  ⚠️  第 {page_num} 頁: 無完整訂單")
                else:
                    print(f"  ⚠️  第 {page_num} 頁: 無法提取文字")
            
            return all_orders
                
    except Exception as e:
        print(f"  ❌ 錯誤: {str(e)}")
        return []

def batch_process_pdfs(input_dir, output_file=None):
    """批次處理PDF檔案"""
    print(f"📁 掃描目錄: {input_dir}")
    
    # 尋找所有PDF檔案
    pdf_files = []
    for file in os.listdir(input_dir):
        if file.lower().endswith('.pdf'):
            pdf_files.append(os.path.join(input_dir, file))
    
    if not pdf_files:
        print("❌ 在指定目錄中未找到PDF檔案")
        return None
    
    print(f"📋 找到 {len(pdf_files)} 個PDF檔案")
    
    # 處理所有PDF檔案
    all_results = []
    successful_files = 0
    
    for pdf_file in pdf_files:
        orders = process_single_pdf(pdf_file)
        if orders:
            all_results.extend(orders)
            successful_files += 1
        print()  # 空行分隔
    
    if all_results:
        # 修正目的地欄位，確保所有記錄都有正確的目的地值
        for order in all_results:
            if not order.get('目的地') or order['目的地'] == '':
                order['目的地'] = 'N/A'
        
        # 建立DataFrame並進行排序
        df = pd.DataFrame(all_results)
        
        # 依照客戶訂單出貨日進行遞增排序
        if not df.empty and '客戶訂單出貨日' in df.columns:
            print("🔄 正在排序所有資料...")
            # 將日期字串轉換為datetime物件進行排序
            # 嘗試多種日期格式
            df['客戶訂單出貨日'] = pd.to_datetime(df['客戶訂單出貨日'], format='%m/%d/%y', errors='coerce')
            df = df.sort_values('客戶訂單出貨日', ascending=True)
            # 將日期轉回字串格式
            df['客戶訂單出貨日'] = df['客戶訂單出貨日'].dt.strftime('%m/%d/%y')
        
        print(f"\n📊 批次處理結果:")
        print(f"✅ 成功處理 {successful_files}/{len(pdf_files)} 個檔案")
        print(f"✅ 總共 {len(all_results)} 筆訂單資料（已按出貨日排序）")
        
        # 顯示統計資訊
        print("\n📈 統計資訊:")
        if '客戶' in df.columns:
            print(f"客戶數量: {df['客戶'].nunique()}")
        if 'PO' in df.columns:
            print(f"PO數量: {df['PO'].nunique()}")
        if '數量' in df.columns:
            print(f"總數量: {df['數量'].sum():,}")
        
        # 顯示前幾筆資料作為預覽
        print("\n📋 資料預覽:")
        preview_cols = ['客戶訂單出貨日', '客戶', 'PO', '數量', '型號', '顏色中文', '來源檔案']
        available_cols = [col for col in preview_cols if col in df.columns]
        print(df[available_cols].head().to_string(index=False))
        
        # 生成Excel檔案
        if output_file is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = f"英文格式批次處理結果_{timestamp}.xlsx"
        
        # 修正目的地欄位，確保N/A值正確顯示
        if '目的地' in df.columns:
            df['目的地'] = df['目的地'].fillna('N/A')
        
        df.to_excel(output_file, index=False, engine='xlsxwriter')
        print(f"\n💾 Excel檔案已生成: {output_file}")
        
        return df
    else:
        print("❌ 未找到任何有效的訂單資料")
        return None

def main():
    parser = argparse.ArgumentParser(description='PDF訂單轉Excel工具 - 英文格式批次處理版本')
    parser.add_argument('input_dir', help='包含PDF檔案的目錄路徑')
    parser.add_argument('-o', '--output', help='輸出Excel檔案路徑（可選）')
    parser.add_argument('--version', action='version', version='PDF訂單轉Excel工具 v2.0 (英文格式版)')
    
    args = parser.parse_args()
    
    if not os.path.isdir(args.input_dir):
        print(f"❌ 錯誤: 目錄不存在 {args.input_dir}")
        sys.exit(1)
    
    print("=" * 60)
    print("PDF訂單轉Excel工具 - 英文格式批次處理版本")
    print("=" * 60)
    
    # 批次處理PDF檔案
    result = batch_process_pdfs(args.input_dir, args.output)
    
    if result is not None:
        print("\n🎉 批次處理完成！")
    else:
        print("\n❌ 批次處理失敗！")
        sys.exit(1)

if __name__ == "__main__":
    main() 
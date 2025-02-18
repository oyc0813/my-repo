import os
import logging
import pandas as pd
import json
from openai import OpenAI

#測試OK

def extract_comments_from_excel(file_path: str, sheet_name: str):
    # 讀取 Excel 檔案
    df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')

    # 提取 'RestaurantName' 和 'Comments' 欄位的第 2 列資料（索引為 1）
    restaurant_index = 16
    restaurant_name = df['RestaurantName'].dropna().iloc[restaurant_index]  # 提取第二行的餐廳名稱
    comment_json = df['Comments'].dropna().iloc[restaurant_index]  # 提取第二行的評論資料

    # 初始化 user_input 字串
    user_input = f"餐廳名稱: {restaurant_name}\n"

    try:
        # 解析 JSON 字串
        comments = json.loads(comment_json)
    except json.JSONDecodeError:
        if isinstance(comment_json, str):
            # 處理可能的格式問題
            items = comment_json.strip('[]').split('}, {')
            items = [item.strip() for item in items if item]  # 確保不處理空項目

            # 嘗試將字串轉換為列表
            comments = []
            for item in items:
                item = item.replace("'", "\"")  # 將單引號替換為雙引號
                try:
                    comments.append(json.loads(f"{{{item}}}"))  # 包裝成字典格式
                except json.JSONDecodeError:
                    continue  # 忽略無法解析的項目

    # 直接使用 JSON 進行分析
    if isinstance(comments, list):
        num_comments = len(comments)
        user_input += f"評論數量: {num_comments}\n"

        # 逐條分析評論內容
        for item in comments:
            commenter_name = item.get('commenter_name')
            comment_star_point = item.get('comment_star_point')
            comment_content = item.get('comment_content')
            user_input += f"評論者: {commenter_name}, 評分: {comment_star_point}, 內容: {comment_content}\n"

    return user_input

def gpt_api_text(prompt: str, user_input: str):
    api_key= os.getenv("gpt_api_key")
    client = OpenAI(api_key=api_key)

    try: 
        completion = client.chat.completions.create(
            temperature=0,
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": prompt
                },
                {
                    "role": "user",
                    "content": user_input  # 將提取的 user_input 用作字符串
                }
            ]
        )
        response_content = completion.choices[0].message.content
        prompt_tokens = completion.usage.prompt_tokens  # 提示 (prompt) token 數
        completion_tokens = completion.usage.completion_tokens  # 回應 (completion) token 數

        # 計算成本（按 1M tokens 計價）
        prompt_cost = (prompt_tokens / 1_000_000) * 0.15  # 0.15 美金 / 1M tokens
        completion_cost = (completion_tokens / 1_000_000) * 0.6  # 0.6 美金 / 1M tokens
        total_cost = prompt_cost + completion_cost

        # 顯示 token 消耗資訊
        print(f"Input Token: {prompt_tokens}")
        print(f"Output Token: {completion_tokens}")
        print(f"總共花費: {total_cost}美金")
        print(f"總共花費: {total_cost * 32.85}台幣")
        
        return response_content
    except Exception as e:
        logging.error(f"API 請求失敗: {e}")
        return "API 請求失敗"

# 使用範例
file_path = 'C:\\Users\\Tibame\\Desktop\\icandoit\\wanhua_restaurant_20250215155839.xlsx'  # 替換為您的文件路徑
sheet_name = 'Sheet1'  # 替換為您的工作表名稱

# 獲取評論資料並生成 user_input
user_input = extract_comments_from_excel(file_path, sheet_name)

# 設定 prompt
prompt = """ 
請幫我分析user輸入的評論在 Google 評論上的內容，並針對以下重點提供摘要：
1. **餐點口味**（評判標準：請以user的內文推薦指數顯示多少%推薦。）(10個字以內)

2. **價位**（評判標準：總體評價、特色、常見評論）

3. **關鍵字**（常見的好評/負評關鍵詞）(10個字以內)  

4. **推薦菜色**（評論中最常被提及的推薦菜品）(10個字以內)   

5. **環境氛圍**(評判標準：請以所有user的內文資訊顯示多少%推薦。)

6. **服務品質**(評判標準：請以所有user的內文資訊顯示多少%推薦。)

7. **適合客群**(評判標準：適合家庭聚會、情侶約會、商務宴請、慶祝場合等，只要給客群類型(10個字以內))

請依照以下範例格式輸出:
   - 餐廳名稱：  
   - 總體評價（1~5分）：  
   - 環境氛圍:
   - 服務品質:
   - 餐點口味： 
   - 價位：
   - 關鍵字： 
   - 推薦菜色：
   - 適合客群:
"""

# 調用 GPT API 進行分析
res = gpt_api_text(prompt=prompt, user_input=user_input)

print(res)

# 確保 res 不為 None 或 "API 請求失敗" 才寫入文件
if res and res != "API 請求失敗":
    with open("google_Comment.md", "w", encoding='utf-8') as f:
        f.write(res)
else:
    print("無法寫入文件，因為 API 請求失敗。")

import requests
from datetime import datetime
from time import sleep
import pandas as pd
from typing import Dict


def parser(product_dict: Dict):
    items = []
    selection_of_goods = []

    for p in product_dict:
        article = str(p.get("id"))
        product_link = f"https://www.wildberries.ru/catalog/{article}/detail.aspx"
        name = p.get("name")
        price = p.get("sizes")[0].get("price").get("product", 0) / 100
        seller_name = p.get("supplier")
        seller_link = f"https://www.wildberries.ru/seller/{p.get('supplierId')}"

        availability_product_sizes = []
        sizes = p.get("sizes")
        for size in sizes:
            availability_product_sizes.append(size.get("origName"))

        rating = p.get("nmReviewRating")
        number_of_reviews = p.get("nmFeedbacks")

        add_url = f"https://kzn-basket-cdn-03bl.geobasket.ru/vol{article[:-5]}/part{article[:-3]}/{article}/info/ru/card.json"
        add_information = s.get(url=add_url).json()

        description = add_information.get("description", "").replace('\n', ' ')

        all_specifications = {}
        specifications = add_information.get("grouped_options", [])
        for spec in specifications:
            for option in spec.get("options"):
                all_specifications[option.get("name")] = option.get("value")

        photo_count = add_information.get("media").get("photo_count", 0)
        image_links = []
        for number in range(1, int(photo_count) + 1):
            url = f"https://kzn-basket-cdn-03bl.geobasket.ru/vol{article[:4]}/part{article[:6]}/{article}/images/big/{number}.webp"
            image_links.append(url)

        size_url = f"https://www.wildberries.ru/__internal/u-card/cards/v4/detail?appType=1&curr=rub&dest=-2133462&spp=30&hide_dtype=9%3B11&ab_testing=false&lang=ru&nm={article}"
        size_information = s.get(url=size_url).json().get("products")[0].get("sizes")

        all_product_sizes = []
        product_stock = 0
        for product_size in size_information:
            all_product_sizes.append(product_size.get("origName"))
            if product_size.get("stocks") != []:
                product_stock += int(product_size.get("stocks")[0].get("qty"))

        data = {
            "product_link": product_link,
            "article": article,
            "name": name,
            "price": price,
            "description": description,
            "image_links": image_links,
            "all_specifications": all_specifications,
            "seller_name": seller_name,
            "seller_link": seller_link,
            "all_product_sizes": all_product_sizes,
            "availability_product_sizes": availability_product_sizes,
            "product_stock": product_stock,
            "rating": rating,
            "number_of_reviews": number_of_reviews,
        }

        items.append(data)

        if rating >= 4.5 and price <= 10000 and all_specifications.get("Страна производства", []) == "Россия":
            selection_of_goods.append(data)

    save_xlsx(items, selection_of_goods)


def save_xlsx(items, selection_of_goods):

    try:
        df_items = pd.DataFrame(items)
        existing_df = pd.read_excel("result.xlsx", sheet_name="wildberries")
        combined_df = pd.concat([existing_df, df_items], ignore_index=True)

        with pd.ExcelWriter("result.xlsx", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            combined_df.to_excel(writer, sheet_name="wildberries", index=False)

    except (FileNotFoundError, ValueError):
        with pd.ExcelWriter("result.xlsx", engine='openpyxl', mode='w') as writer:
            df_items.to_excel(writer, sheet_name="wildberries", index=False)

    try:
        df_select = pd.DataFrame(selection_of_goods)
        existing_df = pd.read_excel("selection_of_goods.xlsx", sheet_name="wildberries")
        combined_df = pd.concat([existing_df, df_select], ignore_index=True)

        with pd.ExcelWriter("selection_of_goods.xlsx", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer_selection:
            combined_df.to_excel(writer_selection, sheet_name="wildberries", index=False)

    except (FileNotFoundError, ValueError):
        with pd.ExcelWriter("selection_of_goods.xlsx", engine='openpyxl', mode='w') as writer_selection:
            df_select.to_excel(writer_selection, sheet_name="wildberries", index=False)


def main():
    page = 1

    while True:
        url = f"https://www.wildberries.ru/__internal/u-search/exactmatch/ru/common/v18/search?ab_testing=false&ab_testing=false&appType=1&curr=rub&dest=-2133462&hide_dtype=9;11&inheritFilters=false&lang=ru&page={page}&query={search_requests}&resultset=catalog&sort=popular&spp=30&suppressSpellcheck=false"
        product = s.get(url=url).json()
        product_dict = product.get("products", [])

        if product_dict == []:
            break

        parser(product_dict)

        if len(product_dict) < 100:
            break

        page += 1
        sleep(0.1)


if __name__ == '__main__':
    search_requests = "пальто из натуральной шерсти"

    headers = {
        'accept': '*/*',
        'accept-language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36',
    }

    cookies = {
        'x_wbaas_token': '1.1000.6aee6eaa9873486796d5d5baa62b6721.MHw5NS4xMDUuNjcuOXxNb3ppbGxhLzUuMCAoV2luZG93cyBOVCAxMC4wOyBXaW42NDsgeDY0KSBBcHBsZVdlYktpdC81MzcuMzYgKEtIVE1MLCBsaWtlIEdlY2tvKSBDaHJvbWUvMTQyLjAuMC4wIFNhZmFyaS81MzcuMzZ8MTc2NTkwNjAyM3xyZXVzYWJsZXwyfGV5Sm9ZWE5vSWpvaUluMD18MHwzfDE3NjUzMDEyMjN8MQ==.MEUCIBVn4GfF2OoIzSuiZylsTbCkvlXaCcyHpEjlb3rhbtOxAiEA2KPMErXli56ssNtMlBketH7THc+glSAsOdcHrMpbCf0=',
    }

    s = requests.Session()

    s.headers.update(headers)
    s.cookies.update(cookies)
    main()


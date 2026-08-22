import pandas as pd
from pathlib import Path
import json


# ================================================================
# 鵡川高校 化石デジタルライブラリー
# data_update.py
#
# Excel → 公開用JSON作成プログラム
#
# 【対応機能】
#
# ① 複数画像対応
# ② 画像ファイルチェック
# ③ サムネイル画像チェック
# ④ 公開用specimens.json生成
# ⑤ 公開用localities_public.json生成
# ⑥ 詳細産地を公開JSONから除外
# ⑦ 採集位置IDを公開JSONから除外
# ⑧ 採集者を公開JSONから除外
# ⑨ 公開用座標のみ公開
# ⑩ 同じ地点IDの複数座標を平均
#
# ================================================================


print("========================================")
print(" 鵡川高校 化石デジタルライブラリー")
print(" データチェック・公開データ作成")
print("========================================")
print()


# ================================================================
# 1. ファイル設定
# ================================================================

SPECIMENS_FILE = "specimens.xlsx"

LOCALITIES_FILE = "localities_open.xlsx"

# 元画像フォルダ
IMAGE_FOLDER = Path("images")

# サムネイルフォルダ
THUMB_FOLDER = IMAGE_FOLDER / "thumbs"

# 公開用JSON
OUTPUT_FOLDER = Path("public_data")


# ================================================================
# 2. Excel読み込み
# ================================================================

print("【1】Excelを読み込んでいます。")
print()


try:

    specimens = pd.read_excel(
        SPECIMENS_FILE,
        dtype=str
    ).fillna("")


    localities = pd.read_excel(
        LOCALITIES_FILE,
        dtype=str
    ).fillna("")


except FileNotFoundError:

    print(
        "ERROR：必要なExcelファイルが見つかりません。"
    )

    print()

    print(
        "以下のファイルが同じフォルダにあるか確認してください。"
    )

    print(
        "  -",
        SPECIMENS_FILE
    )

    print(
        "  -",
        LOCALITIES_FILE
    )

    print()

    raise SystemExit


print(
    "  OK：Excel読み込み完了"
)

print(
    "  標本数：",
    len(specimens)
)

print(
    "  採集位置数：",
    len(localities)
)

print()


# ================================================================
# 3. 列名
# ================================================================

SPECIMEN_ID = "標本Ｎｏ.（MDR〇〇）"

SPECIMEN_POINT_ID = "採集位置ID"

IMAGE = "画像データ"

SITE_ID = "地点ＩＤ"

LATITUDE = "公開用北緯"

LONGITUDE = "公開用東経"


# ================================================================
# 4. 必要な列の確認
# ================================================================

print("【2】Excelの列名を確認しています。")
print()


specimen_required_columns = [

    SPECIMEN_ID,

    "分　　類",

    "学　　名",

    "和　　名",

    "地　層　名",

    "時　　代",

    "採　集　日",

    "備考",

    IMAGE,

    SITE_ID,

    SPECIMEN_POINT_ID

]


locality_required_columns = [

    SITE_ID,

    "採集位置ID",

    "産　　地",

    "地　層　名",

    "時　　代",

    LATITUDE,

    LONGITUDE

]


missing_specimen_columns = [

    column

    for column in specimen_required_columns

    if column not in specimens.columns

]


missing_locality_columns = [

    column

    for column in locality_required_columns

    if column not in localities.columns

]


column_errors = []


for column in missing_specimen_columns:

    column_errors.append(

        f"specimens.xlsx：{column} がありません"

    )


for column in missing_locality_columns:

    column_errors.append(

        f"localities_open.xlsx：{column} がありません"

    )


if len(column_errors) == 0:

    print(
        "  OK：必要な列はすべて存在します。"
    )

else:

    print(
        "  ERROR：必要な列がありません。"
    )

    for error in column_errors:

        print(
            "   -",
            error
        )

    print()

    print(
        "処理を中止します。"
    )

    raise SystemExit


print()


# ================================================================
# 5. 標本番号のチェック
# ================================================================

print("【3】標本番号をチェックしています。")
print()


duplicate_ids = specimens[

    specimens[SPECIMEN_ID].duplicated(
        keep=False
    )

]


empty_specimen_ids = specimens[

    specimens[SPECIMEN_ID]
    .astype(str)
    .str.strip()
    == ""

]


if len(duplicate_ids) == 0:

    print(
        "  OK：標本番号に重複はありません。"
    )

else:

    print(
        "  ERROR：標本番号が重複しています。"
    )

    for value in duplicate_ids[
        SPECIMEN_ID
    ].unique():

        print(
            "   -",
            value
        )


if len(empty_specimen_ids) == 0:

    print(
        "  OK：標本番号の空欄はありません。"
    )

else:

    print(
        "  ERROR：標本番号に空欄があります。"
    )


print()


# ================================================================
# 6. 地点IDの対応チェック
# ================================================================

print("【4】地点IDをチェックしています。")
print()


locality_site_ids = set(

    localities[
        SITE_ID
    ]
    .astype(str)
    .str.strip()

)


site_id_errors = []


for _, row in specimens.iterrows():

    specimen_id = str(
        row[SPECIMEN_ID]
    ).strip()


    site_id = str(
        row[SITE_ID]
    ).strip()


    if site_id == "":

        site_id_errors.append(

            f"{specimen_id}：地点IDが空欄"

        )

    elif site_id not in locality_site_ids:

        site_id_errors.append(

            f"{specimen_id}："
            f"{site_id} がlocalities_open.xlsxにありません"

        )


if len(site_id_errors) == 0:

    print(
        "  OK：すべての標本が地点IDと対応しています。"
    )

else:

    print(
        "  ERROR：地点IDに対応エラーがあります。"
    )

    for error in site_id_errors:

        print(
            "   -",
            error
        )


print()


# ================================================================
# 7. 地点IDの構造確認
# ================================================================

print("【5】地点IDの構造を確認しています。")
print()


site_counts = (

    localities[
        SITE_ID
    ]
    .astype(str)
    .str.strip()
    .value_counts()

)


duplicate_sites = site_counts[
    site_counts > 1
]


if len(duplicate_sites) == 0:

    print(
        "  OK：各地点IDは1件ずつです。"
    )

else:

    print(
        "  INFO：複数の採集位置を持つ地点があります。"
    )

    print(
        "  これは正常なデータとして処理します。"
    )

    for site_id, count in duplicate_sites.items():

        print(
            f"   - {site_id}：{count}件"
        )


print()


# ================================================================
# 8. 複数画像を分割する関数
#
# Excel：
#
# MDR001.jpg;MDR001_02.jpg;MDR001_03.jpg
#
# ↓
#
# [
#   "MDR001.jpg",
#   "MDR001_02.jpg",
#   "MDR001_03.jpg"
# ]
#
# ================================================================

def parse_image_names(value):

    if value is None:

        return []


    text = str(
        value
    ).strip()


    if text == "":

        return []


    # ------------------------------------------------------------
    # セミコロン
    # ------------------------------------------------------------

    text = text.replace(
        "；",
        ";"
    )


    # ------------------------------------------------------------
    # 改行も区切りとして扱う
    # ------------------------------------------------------------

    text = text.replace(
        "\r\n",
        ";"
    )

    text = text.replace(
        "\n",
        ";"
    )


    names = []


    for item in text.split(";"):

        name = item.strip()


        if name == "":

            continue


        names.append(
            name
        )


    return names


# ================================================================
# 9. 画像ファイルのチェック
#
# 元画像：
#
# images/
#
# サムネイル：
#
# images/thumbs/
#
# ================================================================

print("【6】画像ファイルをチェックしています。")
print()


image_errors = []

image_warnings = []


if not IMAGE_FOLDER.exists():

    image_errors.append(
        "imagesフォルダがありません。"
    )

else:

    for _, row in specimens.iterrows():

        specimen_id = str(
            row[SPECIMEN_ID]
        ).strip()


        image_names = parse_image_names(
            row[IMAGE]
        )


        # --------------------------------------------------------
        # 画像名が空欄
        # --------------------------------------------------------

        if len(image_names) == 0:

            image_errors.append(

                f"{specimen_id}：画像名が空欄"

            )

            continue


        # --------------------------------------------------------
        # 複数画像を1枚ずつ確認
        # --------------------------------------------------------

        for image_name in image_names:

            image_path = (

                IMAGE_FOLDER /
                image_name

            )


            # ----------------------------------------------------
            # 元画像
            # ----------------------------------------------------

            if not image_path.exists():

                image_errors.append(

                    f"{specimen_id}："
                    f"{image_name} がありません"

                )

                continue


            # ----------------------------------------------------
            # サムネイル
            # ----------------------------------------------------

            thumbnail_path = (

                THUMB_FOLDER /
                image_name

            )


            if not thumbnail_path.exists():

                image_warnings.append(

                    f"{specimen_id}："
                    f"サムネイル {image_name} がありません"

                )


if len(image_errors) == 0:

    print(
        "  OK：元画像を確認しました。"
    )

else:

    print(
        "  ERROR：画像に問題があります。"
    )

    for error in image_errors:

        print(
            "   -",
            error
        )


if len(image_warnings) > 0:

    print()

    print(
        "  WARNING：サムネイルに確認事項があります。"
    )

    for warning in image_warnings:

        print(
            "   -",
            warning
        )

else:

    print(
        "  OK：サムネイルを確認しました。"
    )


print()


# ================================================================
# 10. 公開用座標のチェック
#
# 同じ地点IDに複数の座標があってもエラーにしない。
#
# 複数の座標は後で平均して代表座標にする。
# ================================================================

print("【7】公開用座標をチェックしています。")
print()


coordinate_errors = []


site_coordinates = {}


for _, row in localities.iterrows():

    site_id = str(
        row[SITE_ID]
    ).strip()


    latitude_text = str(
        row[LATITUDE]
    ).strip()


    longitude_text = str(
        row[LONGITUDE]
    ).strip()


    # ------------------------------------------------------------
    # 北緯
    # ------------------------------------------------------------

    if latitude_text == "":

        coordinate_errors.append(

            f"{site_id}：公開用北緯が空欄"

        )

        continue


    try:

        latitude = float(
            latitude_text
        )

    except ValueError:

        coordinate_errors.append(

            f"{site_id}："
            f"公開用北緯が数値ではありません"

        )

        continue


    # ------------------------------------------------------------
    # 東経
    # ------------------------------------------------------------

    if longitude_text == "":

        coordinate_errors.append(

            f"{site_id}：公開用東経が空欄"

        )

        continue


    try:

        longitude = float(
            longitude_text
        )

    except ValueError:

        coordinate_errors.append(

            f"{site_id}："
            f"公開用東経が数値ではありません"

        )

        continue


    # ------------------------------------------------------------
    # 北海道付近の大まかな範囲
    # ------------------------------------------------------------

    if not (
        40 <= latitude <= 46
    ):

        coordinate_errors.append(

            f"{site_id}："
            f"北緯 {latitude} は要確認"

        )


    if not (
        139 <= longitude <= 146
    ):

        coordinate_errors.append(

            f"{site_id}："
            f"東経 {longitude} は要確認"

        )


    # ------------------------------------------------------------
    # 地点IDごとの座標を保存
    # ------------------------------------------------------------

    if site_id not in site_coordinates:

        site_coordinates[
            site_id
        ] = []


    site_coordinates[
        site_id
    ].append(

        (
            latitude,
            longitude
        )

    )


if len(coordinate_errors) == 0:

    print(
        "  OK：公開用座標を確認しました。"
    )

else:

    print(
        "  ERROR：座標に問題があります。"
    )

    for error in coordinate_errors:

        print(
            "   -",
            error
        )


print()


# ================================================================
# 11. 公開データ仕様の確認
# ================================================================

print("【8】公開データの仕様を確認しています。")
print()


print(
    "  OK：詳細産地はspecimens.jsonに出力しません。"
)

print(
    "  OK：採集位置IDはspecimens.jsonに出力しません。"
)

print(
    "  OK：採集者はspecimens.jsonに出力しません。"
)

print(
    "  OK：複数画像に対応します。"
)

print(
    "  OK：元画像はimages/を使用します。"
)

print(
    "  OK：サムネイルはimages/thumbs/を使用します。"
)

print(
    "  OK：座標はlocalities_public.jsonだけに出力します。"
)

print(
    "  OK：同じ地点IDの複数座標は代表座標にまとめます。"
)

print()


# ================================================================
# 12. 総合チェック結果
# ================================================================

print("========================================")
print(" データチェック結果")
print("========================================")
print()


print(
    "必要列エラー：",
    len(column_errors),
    "件"
)


print(
    "標本番号重複：",
    len(duplicate_ids),
    "件"
)


print(
    "標本番号空欄：",
    len(empty_specimen_ids),
    "件"
)


print(
    "地点ID対応エラー：",
    len(site_id_errors),
    "件"
)


print(
    "画像エラー：",
    len(image_errors),
    "件"
)


print(
    "サムネイル警告：",
    len(image_warnings),
    "件"
)


print(
    "座標エラー：",
    len(coordinate_errors),
    "件"
)


print()


# ================================================================
# 13. エラー判定
#
# 元画像がない場合は公開JSONを作成しない。
#
# サムネイル不足はWARNINGとして扱い、
# JSONは作成する。
# ================================================================

total_errors = (

    len(column_errors)

    +

    len(duplicate_ids)

    +

    len(empty_specimen_ids)

    +

    len(site_id_errors)

    +

    len(image_errors)

    +

    len(coordinate_errors)

)


if total_errors > 0:

    print(
        "========================================"
    )

    print(
        "★★★ データに確認事項があります ★★★"
    )

    print(
        "========================================"
    )

    print()

    print(
        "安全のため、公開用JSONは作成しません。"
    )

    print(
        "エラーを修正してから、もう一度実行してください。"
    )

    raise SystemExit


print(
    "★★★ データチェック正常 ★★★"
)

print()


# ================================================================
# 14. 出力フォルダ作成
# ================================================================

print("【9】公開用JSONを作成しています。")
print()


OUTPUT_FOLDER.mkdir(
    exist_ok=True
)


# ================================================================
# 15. specimens.json
#
# 複数画像対応
#
# Excel：
#
# MDR001.jpg;MDR001_02.jpg
#
# ↓
#
# JSON：
#
# "images": [
#     "images/thumbs/MDR001.jpg",
#     "images/thumbs/MDR001_02.jpg"
# ]
#
# "original_images": [
#     "images/MDR001.jpg",
#     "images/MDR001_02.jpg"
# ]
#
# さらに互換性維持のため、
#
# "image":
#     "images/thumbs/MDR001.jpg"
#
# も残す。
#
# ================================================================

print(
    "  specimens.json を作成しています。"
)


specimen_data = []


for _, row in specimens.iterrows():

    image_names = parse_image_names(
        row[IMAGE]
    )


    # ------------------------------------------------------------
    # サムネイル画像URL
    # ------------------------------------------------------------

    thumbnail_urls = []


    for image_name in image_names:

        thumbnail_path = (

            THUMB_FOLDER /
            image_name

        )


        thumbnail_urls.append(

            thumbnail_path.as_posix()

        )


    # ------------------------------------------------------------
    # 元画像URL
    # ------------------------------------------------------------

    original_urls = []


    for image_name in image_names:

        original_path = (

            IMAGE_FOLDER /
            image_name

        )


        original_urls.append(

            original_path.as_posix()

        )


    # ------------------------------------------------------------
    # 最初の画像
    #
    # 現在のindex.htmlとの互換性維持用
    # ------------------------------------------------------------

    first_image = ""

    if len(thumbnail_urls) > 0:

        first_image = thumbnail_urls[0]


    # ------------------------------------------------------------
    # 公開用データ
    # ------------------------------------------------------------

    data = {

        "specimen_id":
            str(
                row[SPECIMEN_ID]
            ).strip(),


        "classification":
            str(
                row["分　　類"]
            ).strip(),


        "scientific_name":
            str(
                row["学　　名"]
            ).strip(),


        "common_name":
            str(
                row["和　　名"]
            ).strip(),


        "formation":
            str(
                row["地　層　名"]
            ).strip(),


        "age":
            str(
                row["時　　代"]
            ).strip(),


        "collection_date":
            str(
                row["採　集　日"]
            ).strip(),


        "notes":
            str(
                row["備考"]
            ).strip(),


        # --------------------------------------------------------
        # 従来との互換性を維持
        # 最初のサムネイルをimageとして保存
        # --------------------------------------------------------

        "image":
            first_image,


        # --------------------------------------------------------
        # 複数画像
        # サムネイル
        # --------------------------------------------------------

        "images":
            thumbnail_urls,


        # --------------------------------------------------------
        # 複数画像
        # 元画像
        # --------------------------------------------------------

        "original_images":
            original_urls,


        # --------------------------------------------------------
        # 地点ID
        # --------------------------------------------------------

        "site_id":
            str(
                row[SITE_ID]
            ).strip()

    }


    specimen_data.append(
        data
    )


with open(

    OUTPUT_FOLDER /
    "specimens.json",

    "w",

    encoding="utf-8"

) as f:

    json.dump(

        specimen_data,

        f,

        ensure_ascii=False,

        indent=2

    )


print(
    "  OK：specimens.jsonを作成しました。"
)

print()


# ================================================================
# 16. localities_public.json
#
# 同じ地点IDの座標を平均する。
#
# 詳細な採集位置IDは公開しない。
# ================================================================

print(
    "  localities_public.json を作成しています。"
)


locality_data = []


processed_site_ids = set()


for _, row in localities.iterrows():

    site_id = str(
        row[SITE_ID]
    ).strip()


    # ------------------------------------------------------------
    # すでに処理した地点IDはスキップ
    # ------------------------------------------------------------

    if site_id in processed_site_ids:

        continue


    processed_site_ids.add(
        site_id
    )


    coordinates = site_coordinates.get(
        site_id,
        []
    )


    if len(coordinates) == 0:

        continue


    # ------------------------------------------------------------
    # 平均緯度
    # ------------------------------------------------------------

    average_latitude = (

        sum(

            coordinate[0]

            for coordinate in coordinates

        )

        /

        len(coordinates)

    )


    # ------------------------------------------------------------
    # 平均経度
    # ------------------------------------------------------------

    average_longitude = (

        sum(

            coordinate[1]

            for coordinate in coordinates

        )

        /

        len(coordinates)

    )


    # ------------------------------------------------------------
    # 公開用地点データ
    #
    # 採集位置IDは出力しない。
    # ------------------------------------------------------------

    data = {

        "site_id":
            site_id,


        "place":
            str(
                row["産　　地"]
            ).strip(),


        "formation":
            str(
                row["地　層　名"]
            ).strip(),


        "age":
            str(
                row["時　　代"]
            ).strip(),


        "latitude":
            average_latitude,


        "longitude":
            average_longitude

    }


    locality_data.append(
        data
    )


with open(

    OUTPUT_FOLDER /
    "localities_public.json",

    "w",

    encoding="utf-8"

) as f:

    json.dump(

        locality_data,

        f,

        ensure_ascii=False,

        indent=2

    )


print(
    "  OK：localities_public.jsonを作成しました。"
)

print()


# ================================================================
# 17. 最終確認
# ================================================================

print("========================================")
print(" 公開用データ作成完了")
print("========================================")
print()


print(
    "作成されたファイル："
)


print(
    "  - public_data/specimens.json"
)


print(
    "  - public_data/localities_public.json"
)


print()


print(
    "標本データ数：",
    len(specimen_data)
)


print(
    "公開地点数：",
    len(locality_data)
)


print()


print(
    "画像設定："
)


print(
    "  元画像：images/"
)


print(
    "  サムネイル：images/thumbs/"
)


print(
    "  複数画像：Excelの「画像データ」を ; で区切る"
)


print()


print(
    "公開JSONの画像データ："
)


print(
    "  image → 最初のサムネイル"
)


print(
    "  images → 全サムネイル"
)


print(
    "  original_images → 全元画像"
)


print()


print(
    "公開データの仕様："
)


print(
    "  ① 詳細産地 → specimens.jsonから除外"
)


print(
    "  ② 採集位置ID → specimens.jsonから除外"
)


print(
    "  ③ 採集者 → specimens.jsonから除外"
)


print(
    "  ④ 画像 → images/を使用"
)


print(
    "  ⑤ サムネイル → images/thumbs/を使用"
)


print(
    "  ⑥ 複数画像 → images配列として公開"
)


print(
    "  ⑦ 元画像 → original_images配列として公開"
)


print(
    "  ⑧ 公開用座標 → localities_public.jsonのみ"
)


print(
    "  ⑨ 同じ地点IDの座標 → 平均して代表座標に統合"
)


print()


print(
    "処理終了"
)
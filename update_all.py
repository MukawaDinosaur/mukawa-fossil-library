import pandas as pd
from pathlib import Path
from PIL import Image
import json
import shutil
import sys


# ================================================================
# 鵡川高校 化石デジタルライブラリー
# update_all.py
#
# ================================================================
# このプログラムで行う処理
#
# ① specimens.xlsx 読み込み
# ② localities_open.xlsx 読み込み
# ③ 標本番号チェック
# ④ 地点IDチェック
# ⑤ 画像チェック
# ⑥ 元画像からサムネイルを自動生成
# ⑦ specimens.json 作成
# ⑧ localities_public.json 作成
#
# ================================================================
#
# フォルダ構成
#
# 鵡川高校＿化石デジタルライブラリー/
#
# ├─ update_all.py
# ├─ specimens.xlsx
# ├─ localities_open.xlsx
# │
# ├─ images/
# │   ├─ MDR260516-01.jpg
# │   ├─ MDR260516-02.jpg
# │   └─ thumbs/
# │       ├─ MDR260516-01.jpg
# │       └─ MDR260516-02.jpg
# │
# └─ public_data/
#     ├─ specimens.json
#     └─ localities_public.json
#
# ================================================================


print()
print("============================================================")
print(" 鵡川高校 化石デジタルライブラリー")
print(" データ更新プログラム")
print(" update_all.py")
print("============================================================")
print()


# ================================================================
# 1. ファイル・フォルダ設定
# ================================================================

SPECIMENS_FILE = Path("specimens.xlsx")

LOCALITIES_FILE = Path("localities_open.xlsx")

IMAGE_FOLDER = Path("images")

THUMB_FOLDER = IMAGE_FOLDER / "thumbs"

OUTPUT_FOLDER = Path("public_data")


# ================================================================
# サムネイル設定
# ================================================================

# サムネイルの長辺
THUMB_MAX_SIZE = 1200

# JPEG品質
THUMB_QUALITY = 88


# ================================================================
# Excel列名
# ================================================================

SPECIMEN_ID = "標本Ｎｏ.（MDR〇〇）"

SPECIMEN_POINT_ID = "採集位置ID"

IMAGE = "画像データ"

SITE_ID = "地点ＩＤ"

LATITUDE = "公開用北緯"

LONGITUDE = "公開用東経"


# ================================================================
# 2. 必要ファイル確認
# ================================================================

print("【1】必要ファイルを確認しています。")
print()


required_files = [

    SPECIMENS_FILE,

    LOCALITIES_FILE

]


file_errors = []


for file_path in required_files:

    if not file_path.exists():

        file_errors.append(
            f"{file_path} がありません"
        )


if len(file_errors) > 0:

    print("ERROR：必要なファイルがありません。")
    print()

    for error in file_errors:

        print(" -", error)

    print()

    print("処理を中止します。")

    sys.exit(1)


print("  OK：必要ファイルを確認しました。")
print()


# ================================================================
# 3. Excel読み込み
# ================================================================

print("【2】Excelを読み込んでいます。")
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


except Exception as error:

    print("ERROR：Excelの読み込みに失敗しました。")
    print()
    print(error)
    print()

    sys.exit(1)


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
# 4. 必要列確認
# ================================================================

print("【3】Excelの列名を確認しています。")
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

    sys.exit(1)


print()


# ================================================================
# 5. 標本番号チェック
# ================================================================

print("【4】標本番号をチェックしています。")
print()


duplicate_ids = specimens[

    specimens[
        SPECIMEN_ID
    ].duplicated(
        keep=False
    )

]


empty_specimen_ids = specimens[

    specimens[
        SPECIMEN_ID
    ].str.strip() == ""

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
# 6. 地点IDチェック
# ================================================================

print("【5】地点IDをチェックしています。")
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
# 7. 地点ID構造確認
# ================================================================

print("【6】地点IDの構造を確認しています。")
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
# 8. 画像ファイルチェック
# ================================================================

print("【7】画像ファイルをチェックしています。")
print()


image_errors = []

image_paths = []


if not IMAGE_FOLDER.exists():

    image_errors.append(
        "imagesフォルダがありません。"
    )

else:

    for _, row in specimens.iterrows():

        specimen_id = str(
            row[SPECIMEN_ID]
        ).strip()


        image_text = str(
            row[IMAGE]
        ).strip()


        # --------------------------------------------------------
        # 複数画像対応
        #
        # Excelには
        #
        # MDR260516-01.jpg
        #
        # または
        #
        # MDR260516-01.jpg;MDR260516-01_2.jpg
        #
        # のように記載可能。
        #
        # --------------------------------------------------------

        if image_text == "":

            image_errors.append(

                f"{specimen_id}：画像名が空欄"

            )

            continue


        image_names = [

            name.strip()

            for name in image_text.split(";")

            if name.strip() != ""

        ]


        specimen_image_paths = []


        for image_name in image_names:

            image_path = (

                IMAGE_FOLDER /

                image_name

            )


            if not image_path.exists():

                image_errors.append(

                    f"{specimen_id}："
                    f"{image_name} がありません"

                )

            else:

                specimen_image_paths.append(
                    image_path
                )


        image_paths.append(

            (
                specimen_id,
                specimen_image_paths
            )

        )


if len(image_errors) == 0:

    print(
        "  OK：画像ファイルを確認しました。"
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


print()


# ================================================================
# 9. 公開用座標チェック
# ================================================================

print("【8】公開用座標をチェックしています。")
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


    if site_id not in site_coordinates:

        site_coordinates[site_id] = []


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
# 10. エラー総合判定
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
    "座標エラー：",
    len(coordinate_errors),
    "件"
)


print()


if total_errors > 0:

    print(
        "★★★ データに確認事項があります ★★★"
    )

    print()

    print(
        "安全のため、"
        "JSON作成・サムネイル作成は行いません。"
    )

    print()

    print(
        "エラーを修正してから"
        "もう一度 update_all.py を実行してください。"
    )

    sys.exit(1)


print(
    "★★★ データチェック正常 ★★★"
)

print()


# ================================================================
# 11. 出力フォルダ作成
# ================================================================

print("【9】出力フォルダを準備しています。")
print()


OUTPUT_FOLDER.mkdir(
    exist_ok=True
)


THUMB_FOLDER.mkdir(
    exist_ok=True
)


print(
    "  OK：public_data/ を確認しました。"
)


print(
    "  OK：images/thumbs/ を確認しました。"
)

print()


# ================================================================
# 12. サムネイル作成
# ================================================================

print("【10】サムネイルを作成しています。")
print()


thumbnail_errors = []

thumbnail_created = 0

thumbnail_skipped = 0


# ---------------------------------------------------------------
# 対象画像をimagesフォルダから取得
#
# thumbsフォルダ自身は除外する。
# ---------------------------------------------------------------

supported_extensions = {

    ".jpg",
    ".jpeg",
    ".png",
    ".webp"

}


source_images = [

    path

    for path in IMAGE_FOLDER.iterdir()

    if path.is_file()

    and path.suffix.lower()
    in supported_extensions

]


print(
    "  対象画像：",
    len(source_images),
    "枚"
)

print()


for source_path in source_images:

    try:

        thumbnail_path = (

            THUMB_FOLDER /

            source_path.name

        )


        # --------------------------------------------------------
        # すでに同じ画像からサムネイルが存在する場合
        #
        # 元画像の更新日時より新しければスキップ
        # --------------------------------------------------------

        if thumbnail_path.exists():

            source_mtime = (
                source_path.stat().st_mtime
            )

            thumb_mtime = (
                thumbnail_path.stat().st_mtime
            )


            if thumb_mtime >= source_mtime:

                thumbnail_skipped += 1

                continue


        # --------------------------------------------------------
        # 画像を開く
        # --------------------------------------------------------

        with Image.open(
            source_path
        ) as img:

            # ----------------------------------------------------
            # EXIF回転を反映
            # ----------------------------------------------------

            try:

                from PIL import ImageOps

                img = ImageOps.exif_transpose(
                    img
                )

            except Exception:

                pass


            # ----------------------------------------------------
            # RGBへ変換
            #
            # JPEG保存時のエラー防止
            # ----------------------------------------------------

            if img.mode not in (
                "RGB",
                "L"
            ):

                if "A" in img.getbands():

                    background = Image.new(
                        "RGB",
                        img.size,
                        "white"
                    )

                    background.paste(
                        img,
                        mask=img.getchannel("A")
                    )

                    img = background

                else:

                    img = img.convert(
                        "RGB"
                    )


            else:

                if img.mode == "L":

                    img = img.convert(
                        "RGB"
                    )


            # ----------------------------------------------------
            # 長辺をTHUMB_MAX_SIZE以下にする
            # ----------------------------------------------------

            img.thumbnail(

                (
                    THUMB_MAX_SIZE,
                    THUMB_MAX_SIZE
                ),

                Image.Resampling.LANCZOS

            )


            # ----------------------------------------------------
            # 保存
            # ----------------------------------------------------

            # PNGはPNGとして保存
            if source_path.suffix.lower() == ".png":

                img.save(

                    thumbnail_path,

                    "PNG",

                    optimize=True

                )

            else:

                # JPEGとして保存
                #
                # 元画像の拡張子がjpg/jpegの場合
                # 同じ拡張子で保存する。
                #
                img.save(

                    thumbnail_path,

                    "JPEG",

                    quality=THUMB_QUALITY,

                    optimize=True

                )


        thumbnail_created += 1


        print(
            "  作成：",
            thumbnail_path
        )


    except Exception as error:

        thumbnail_errors.append(

            f"{source_path.name}：{error}"

        )


if len(thumbnail_errors) > 0:

    print()

    print(
        "ERROR：サムネイル作成中に問題が発生しました。"
    )

    for error in thumbnail_errors:

        print(
            "   -",
            error
        )

    print()

    print(
        "安全のため、JSON作成を中止します。"
    )

    sys.exit(1)


print()

print(
    "  新規・更新：",
    thumbnail_created,
    "枚"
)


print(
    "  変更なし：",
    thumbnail_skipped,
    "枚"
)


print()


# ================================================================
# 13. specimens.json
# ================================================================

print("【11】specimens.jsonを作成しています。")
print()


specimen_data = []


for _, row in specimens.iterrows():

    specimen_id = str(
        row[SPECIMEN_ID]
    ).strip()


    image_text = str(
        row[IMAGE]
    ).strip()


    # ------------------------------------------------------------
    # 複数画像
    #
    # Excel：
    #
    # MDR001.jpg;MDR001_2.jpg;MDR001_3.jpg
    #
    # JSON：
    #
    # "images": [
    #     "images/thumbs/MDR001.jpg",
    #     "images/thumbs/MDR001_2.jpg",
    #     "images/thumbs/MDR001_3.jpg"
    # ]
    #
    # ------------------------------------------------------------

    image_names = [

        name.strip()

        for name in image_text.split(";")

        if name.strip() != ""

    ]


    thumbnail_urls = []

    original_urls = []


    for image_name in image_names:

        original_path = (

            IMAGE_FOLDER /

            image_name

        )


        thumbnail_path = (

            THUMB_FOLDER /

            image_name

        )


        # Web用パス
        original_url = (
            original_path.as_posix()
        )


        thumbnail_url = (
            thumbnail_path.as_posix()
        )


        original_urls.append(
            original_url
        )


        thumbnail_urls.append(
            thumbnail_url
        )


    # ------------------------------------------------------------
    # 後方互換用
    #
    # index.htmlを新仕様へ変更するまで
    # imageにも1枚目を入れておく。
    # ------------------------------------------------------------

    first_original_image = ""

    first_thumbnail_image = ""


    if len(original_urls) > 0:

        first_original_image = (
            original_urls[0]
        )


    if len(thumbnail_urls) > 0:

        first_thumbnail_image = (
            thumbnail_urls[0]
        )


    # ------------------------------------------------------------
    # データ作成
    # ------------------------------------------------------------

    data = {

        "specimen_id":
            specimen_id,


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
        # 旧仕様
        # --------------------------------------------------------

        "image":
            first_original_image,


        # --------------------------------------------------------
        # 新仕様
        # --------------------------------------------------------

        "thumbnail":
            first_thumbnail_image,


        "images":
            thumbnail_urls,


        "original_images":
            original_urls,


        # --------------------------------------------------------
        # 地点ID
        #
        # 詳細な採集位置IDは公開しない。
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


print(
    "  標本数：",
    len(specimen_data)
)

print()


# ================================================================
# 14. localities_public.json
# ================================================================

print(
    "【12】localities_public.jsonを作成しています。"
)

print()


locality_data = []


processed_site_ids = set()


for _, row in localities.iterrows():

    site_id = str(
        row[SITE_ID]
    ).strip()


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


print(
    "  公開地点数：",
    len(locality_data)
)

print()


# ================================================================
# 15. 作成データ確認
# ================================================================

print("【13】作成されたデータを確認しています。")
print()


# specimens.json確認
specimens_json = (
    OUTPUT_FOLDER /
    "specimens.json"
)


if specimens_json.exists():

    print(
        "  OK：specimens.json"
    )

else:

    print(
        "  ERROR：specimens.jsonがありません。"
    )


# localities_public.json確認
localities_json = (
    OUTPUT_FOLDER /
    "localities_public.json"
)


if localities_json.exists():

    print(
        "  OK：localities_public.json"
    )

else:

    print(
        "  ERROR：localities_public.jsonがありません。"
    )


print()


# ================================================================
# 16. 最終結果
# ================================================================

print()
print("============================================================")
print(" 更新処理が完了しました")
print("============================================================")
print()


print("【更新内容】")
print()


print(
    "① Excel読み込み"
)

print(
    "② データチェック"
)

print(
    "③ 画像チェック"
)

print(
    "④ サムネイル生成"
)

print(
    "⑤ specimens.json更新"
)

print(
    "⑥ localities_public.json更新"
)

print()


print("【作成・更新されたファイル】")
print()


print(
    "  images/thumbs/"
)

print(
    "  public_data/specimens.json"
)

print(
    "  public_data/localities_public.json"
)

print()


print("【画像仕様】")
print()


print(
    f"  サムネイル長辺：最大 {THUMB_MAX_SIZE}px"
)

print(
    f"  JPEG品質：{THUMB_QUALITY}"
)

print()


print("【複数画像について】")
print()


print(
    "Excelの「画像データ」欄に"
)

print(
    "画像ファイル名を「;」で区切って入力できます。"
)

print()


print(
    "例："
)

print(
    "MDR260516-01.jpg;"
    "MDR260516-01_2.jpg;"
    "MDR260516-01_3.jpg"
)

print()


print("【公開データの仕様】")
print()


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
    "  ④ サムネイル → images/thumbs/"
)

print(
    "  ⑤ 元画像 → images/"
)

print(
    "  ⑥ 公開用座標 → localities_public.json"
)

print(
    "  ⑦ 同じ地点IDの複数座標 → 平均して代表座標に統合"
)

print()


print("処理終了。")
print()
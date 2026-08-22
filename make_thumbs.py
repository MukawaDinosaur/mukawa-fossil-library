from pathlib import Path
from PIL import Image

# ============================================================
# 鵡川高校 化石デジタルライブラリー
# サムネイル自動生成プログラム
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

IMAGE_DIR = BASE_DIR / "images"

THUMB_DIR = IMAGE_DIR / "thumbs"


# ============================================================
# 設定
# ============================================================

MAX_SIZE = 1200

JPEG_QUALITY = 85


IMAGE_EXTENSIONS = {
    ".jpg",
    ".jpeg",
    ".png",
    ".webp"
}


# ============================================================
# 開始
# ============================================================

print("================================")
print("鵡川高校 化石デジタルライブラリー")
print("サムネイル自動生成")
print("================================")
print()


# ============================================================
# imagesフォルダ確認
# ============================================================

if not IMAGE_DIR.exists():

    print("ERROR：imagesフォルダが見つかりません。")

    print()
    print("現在の場所：")
    print(BASE_DIR)

    print()

    input("Enterキーを押して終了します。")

    raise SystemExit


# ============================================================
# thumbsフォルダ作成
# ============================================================

THUMB_DIR.mkdir(
    parents=True,
    exist_ok=True
)


print("元画像フォルダ：")
print(IMAGE_DIR)

print()

print("サムネイル保存先：")
print(THUMB_DIR)

print()


# ============================================================
# 対象画像を検索
# ============================================================

images = []

for path in IMAGE_DIR.iterdir():

    # thumbsフォルダなどのフォルダは除外
    if path.is_dir():
        continue

    if path.suffix.lower() in IMAGE_EXTENSIONS:
        images.append(path)


# ============================================================
# 画像がない場合
# ============================================================

if len(images) == 0:

    print("画像が見つかりませんでした。")

    print()

    input("Enterキーを押して終了します。")

    raise SystemExit


print(
    f"対象画像：{len(images)}枚"
)

print()


# ============================================================
# カウンター
# ============================================================

success_count = 0

skip_count = 0

error_count = 0


# ============================================================
# サムネイル生成
# ============================================================

for image_path in images:

    print(
        f"処理中：{image_path.name}"
    )


    # --------------------------------------------------------
    # 出力ファイル
    # --------------------------------------------------------

    output_path = (
        THUMB_DIR
        /
        f"{image_path.stem}.jpg"
    )


    try:

        # ----------------------------------------------------
        # 画像を開く
        # ----------------------------------------------------

        with Image.open(
            image_path
        ) as original:


            # ------------------------------------------------
            # RGBへ変換
            # ------------------------------------------------

            if original.mode != "RGB":

                img = original.convert(
                    "RGB"
                )

            else:

                img = original.copy()


            # ------------------------------------------------
            # 元サイズ
            # ------------------------------------------------

            original_width, original_height = img.size


            # ------------------------------------------------
            # すでに小さい画像は拡大しない
            # ------------------------------------------------

            if (
                original_width <= MAX_SIZE
                and
                original_height <= MAX_SIZE
            ):

                resized = img


            else:

                resized = img.copy()

                resized.thumbnail(
                    (
                        MAX_SIZE,
                        MAX_SIZE
                    ),
                    Image.Resampling.LANCZOS
                )


            # ------------------------------------------------
            # JPEG保存
            # ------------------------------------------------

            resized.save(
                output_path,
                "JPEG",
                quality=JPEG_QUALITY,
                optimize=True
            )


            # ------------------------------------------------
            # ファイル容量
            # ------------------------------------------------

            original_size = (
                image_path.stat().st_size
            )

            thumbnail_size = (
                output_path.stat().st_size
            )


            original_mb = (
                original_size
                /
                1024
                /
                1024
            )


            thumbnail_mb = (
                thumbnail_size
                /
                1024
                /
                1024
            )


            print("  OK")

            print(
                f"  元画像："
                f"{original_width} × "
                f"{original_height}px"
            )

            print(
                f"  サムネイル："
                f"{resized.width} × "
                f"{resized.height}px"
            )

            print(
                f"  容量："
                f"{original_mb:.2f} MB → "
                f"{thumbnail_mb:.2f} MB"
            )

            print()


            success_count += 1


    except Exception as e:

        print(
            f"  ERROR：{e}"
        )

        print()

        error_count += 1


# ============================================================
# 結果
# ============================================================

print("================================")
print("サムネイル生成結果")
print("================================")

print()

print(
    f"対象画像：{len(images)}枚"
)

print(
    f"生成成功：{success_count}枚"
)

print(
    f"スキップ：{skip_count}枚"
)

print(
    f"エラー：{error_count}枚"
)

print()

print("保存先：")

print(
    THUMB_DIR
)

print()


if error_count == 0:

    print(
        "★★★ サムネイル生成正常終了 ★★★"
    )

else:

    print(
        "★★★ 一部画像でエラーがあります ★★★"
    )


print()

input(
    "Enterキーを押して終了します。"
)
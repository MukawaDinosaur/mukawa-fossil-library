```python
import pandas as pd
from pathlib import Path
from PIL import Image
import json
import subprocess
import sys
import os


# ================================================================
# 鵡川高校 化石デジタルライブラリー
#
# update_all.py
#
# Excel・画像を追加した後、
# このファイルを実行するだけで
#
# ① データチェック
# ② 公開JSON生成
# ③ サムネイル生成
# ④ Gitへ登録
# ⑤ GitHubへpush
#
# まで自動実行する。
#
# ================================================================


# ================================================================
# 基本設定
# ================================================================

SPECIMENS_FILE = "specimens.xlsx"

LOCALITIES_FILE = "localities_open.xlsx"

IMAGE_FOLDER = Path("images")

THUMB_FOLDER = IMAGE_FOLDER / "thumbs"

OUTPUT_FOLDER = Path("public_data")


# ================================================================
# サムネイル設定
# ================================================================

THUMB_MAX_SIZE = (1200, 1200)

THUMB_QUALITY = 85


# ================================================================
# Git設定
# ================================================================

GIT_REMOTE = "origin"

GIT_BRANCH = "main"


# ================================================================
# 表示
# ================================================================

def print_title(text):

    print()
    print("=" * 60)
    print(text)
    print("=" * 60)
    print()


def print_ok(text):

    print("  OK：", text)


def print_error(text):

    print("  ERROR：", text)


def print_info(text):

    print("  INFO：", text)


# ================================================================
# Gitコマンド実行
# ================================================================

def run_git_command(args):

    try:

        result = subprocess.run(
            args,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace"
        )

    except FileNotFoundError:

        print_error(
            "Gitが見つかりません。"
        )

        print()

        print(
            "Git for Windowsがインストールされているか確認してください。"
        )

        return False, ""

    if result.stdout:

        print(result.stdout)

    if result.stderr:

        print(result.stderr)

    if result.returncode != 0:

        return False, result.stdout + result.stderr

    return True, result.stdout + result.stderr


# ================================================================
# Gitリポジトリ確認
# ================================================================

def check_git_repository():

    print_title("【Git】Gitリポジトリを確認しています")

    success, output = run_git_command(
        [
            "git",
            "rev-parse",
            "--is-inside-work-tree"
        ]
    )

    if not success:

        print_error(
            "このフォルダはGitリポジトリではありません。"
        )

        return False

    print_ok(
        "Gitリポジトリを確認しました。"
    )

    return True


# ================================================================
# Excel読み込み
# ================================================================

def load_excel():

    print_title("【1】Excelを読み込んでいます")

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

        print_error(
            "必要なExcelファイルが見つかりません。"
        )

        print()

        print(
            "必要なファイル："
        )

        print(
            "  -",
            SPECIMENS_FILE
        )

        print(
            "  -",
            LOCALITIES_FILE
        )

        return None, None

    except Exception as error:

        print_error(
            f"Excel読み込み中にエラーが発生しました：{error}"
        )

        return None, None

    print_ok(
        "Excel読み込み完了"
    )

    print(
        "  標本数：",
        len(specimens)
    )

    print(
        "  採集位置数：",
        len(localities)
    )

    return specimens, localities


# ================================================================
# 列名
# ================================================================

SPECIMEN_ID = "標本Ｎｏ.（MDR〇〇）"

SPECIMEN_POINT_ID = "採集位置ID"

IMAGE = "画像データ"

SITE_ID = "地点ＩＤ"

LATITUDE = "公開用北緯"

LONGITUDE = "公開用東経"


# ================================================================
# 必要列チェック
# ================================================================

def check_columns(
    specimens,
    localities
):

    print_title("【2】Excelの列名を確認しています")

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

    errors = []

    for column in specimen_required_columns:

        if column not in specimens.columns:

            errors.append(
                f"specimens.xlsx：{column} がありません"
            )

    for column in locality_required_columns:

        if column not in localities.columns:

            errors.append(
                f"localities_open.xlsx：{column} がありません"
            )

    if errors:

        for error in errors:

            print_error(error)

        return errors

    print_ok(
        "必要な列はすべて存在します。"
    )

    return []


# ================================================================
# 標本番号チェック
# ================================================================

def check_specimen_ids(specimens):

    print_title("【3】標本番号をチェックしています")

    errors = []

    duplicate_ids = specimens[
        specimens[SPECIMEN_ID].duplicated(
            keep=False
        )
    ]

    empty_ids = specimens[
        specimens[SPECIMEN_ID]
        .astype(str)
        .str.strip()
        == ""
    ]

    if len(duplicate_ids) > 0:

        for value in duplicate_ids[
            SPECIMEN_ID
        ].unique():

            errors.append(
                f"標本番号重複：{value}"
            )

    if len(empty_ids) > 0:

        for _, row in empty_ids.iterrows():

            errors.append(
                "標本番号が空欄です。"
            )

    if errors:

        for error in errors:

            print_error(error)

    else:

        print_ok(
            "標本番号に問題ありません。"
        )

    return errors


# ================================================================
# 地点IDチェック
# ================================================================

def check_site_ids(
    specimens,
    localities
):

    print_title("【4】地点IDをチェックしています")

    errors = []

    locality_site_ids = set(

        localities[
            SITE_ID
        ]
        .astype(str)
        .str.strip()

    )

    for _, row in specimens.iterrows():

        specimen_id = str(
            row[SPECIMEN_ID]
        ).strip()

        site_id = str(
            row[SITE_ID]
        ).strip()

        if site_id == "":

            errors.append(
                f"{specimen_id}：地点IDが空欄"
            )

        elif site_id not in locality_site_ids:

            errors.append(
                f"{specimen_id}："
                f"{site_id} がlocalities_open.xlsxにありません"
            )

    if errors:

        for error in errors:

            print_error(error)

    else:

        print_ok(
            "すべての標本が地点IDと対応しています。"
        )

    return errors


# ================================================================
# 複数画像の取得
#
# Excelの「画像データ」欄に
#
# MDR260516-01-01.jpg
# MDR260516-01-02.jpg
# MDR260516-01-03.jpg
#
# のように「,」で複数指定できる。
#
# または
#
# MDR260516-01-01.jpg;MDR260516-01-02.jpg
#
# にも対応。
# ================================================================

def get_image_names(value):

    if value is None:

        return []

    text = str(value).strip()

    if text == "":

        return []

    text = text.replace(
        "；",
        ";"
    )

    text = text.replace(
        "、",
        ","
    )

    text = text.replace(
        "\n",
        ","
    )

    text = text.replace(
        "\r",
        ","
    )

    names = []

    for part in text.replace(
        ";",
        ","
    ).split(","):

        name = part.strip()

        if name:

            names.append(name)

    return names


# ================================================================
# 画像チェック
# ================================================================

def check_images(specimens):

    print_title("【5】画像ファイルをチェックしています")

    errors = []

    if not IMAGE_FOLDER.exists():

        errors.append(
            "imagesフォルダがありません。"
        )

        return errors

    for _, row in specimens.iterrows():

        specimen_id = str(
            row[SPECIMEN_ID]
        ).strip()

        image_names = get_image_names(
            row[IMAGE]
        )

        if len(image_names) == 0:

            errors.append(
                f"{specimen_id}：画像名が空欄"
            )

            continue

        for image_name in image_names:

            image_path = (
                IMAGE_FOLDER /
                image_name
            )

            if not image_path.exists():

                errors.append(
                    f"{specimen_id}："
                    f"{image_name} がありません"
                )

    if errors:

        for error in errors:

            print_error(error)

    else:

        print_ok(
            "画像ファイルをすべて確認しました。"
        )

    return errors


# ================================================================
# 座標チェック
# ================================================================

def check_coordinates(localities):

    print_title("【6】公開用座標をチェックしています")

    errors = []

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

            errors.append(
                f"{site_id}：公開用北緯が空欄"
            )

            continue

        if longitude_text == "":

            errors.append(
                f"{site_id}：公開用東経が空欄"
            )

            continue

        try:

            latitude = float(
                latitude_text
            )

            longitude = float(
                longitude_text
            )

        except ValueError:

            errors.append(
                f"{site_id}：座標が数値ではありません"
            )

            continue

        if not (
            40 <= latitude <= 46
        ):

            errors.append(
                f"{site_id}："
                f"北緯 {latitude} は要確認"
            )

        if not (
            139 <= longitude <= 146
        ):

            errors.append(
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

    if errors:

        for error in errors:

            print_error(error)

    else:

        print_ok(
            "公開用座標を確認しました。"
        )

    return errors, site_coordinates


# ================================================================
# サムネイル生成
# ================================================================

def create_thumbnails():

    print_title("【7】サムネイルを作成しています")

    if not IMAGE_FOLDER.exists():

        print_error(
            "imagesフォルダがありません。"
        )

        return False

    THUMB_FOLDER.mkdir(
        parents=True,
        exist_ok=True
    )

    image_extensions = {

        ".jpg",
        ".jpeg",
        ".png",
        ".webp"

    }

    image_files = []

    for path in IMAGE_FOLDER.rglob("*"):

        if not path.is_file():

            continue

        if THUMB_FOLDER in path.parents:

            continue

        if path.suffix.lower() not in image_extensions:

            continue

        image_files.append(path)

    created = 0

    skipped = 0

    errors = 0

    for image_path in image_files:

        try:

            relative_path = image_path.relative_to(
                IMAGE_FOLDER
            )

            thumb_path = (
                THUMB_FOLDER /
                relative_path
            )

            thumb_path.parent.mkdir(
                parents=True,
                exist_ok=True
            )

            if (
                thumb_path.exists()
                and
                thumb_path.stat().st_mtime
                >=
                image_path.stat().st_mtime
            ):

                skipped += 1

                continue

            with Image.open(
                image_path
            ) as image:

                image.thumbnail(
                    THUMB_MAX_SIZE,
                    Image.Resampling.LANCZOS
                )

                if image.mode in (
                    "RGBA",
                    "P"
                ):

                    image = image.convert(
                        "RGB"
                    )

                image.save(
                    thumb_path,
                    quality=THUMB_QUALITY,
                    optimize=True
                )

            created += 1

            print(
                "  作成：",
                thumb_path
            )

        except Exception as error:

            errors += 1

            print_error(
                f"{image_path}：{error}"
            )

    print()

    print(
        "  新規・更新：",
        created,
        "件"
    )

    print(
        "  変更なし：",
        skipped,
        "件"
    )

    print(
        "  エラー：",
        errors,
        "件"
    )

    return errors == 0


# ================================================================
# 公開JSON生成
# ================================================================

def create_public_json(
    specimens,
    localities,
    site_coordinates
):

    print_title("【8】公開用JSONを作成しています")

    OUTPUT_FOLDER.mkdir(
        parents=True,
        exist_ok=True
    )

    specimen_data = []

    for _, row in specimens.iterrows():

        image_names = get_image_names(
            row[IMAGE]
        )

        image_urls = []

        for image_name in image_names:

            original_path = (
                IMAGE_FOLDER /
                image_name
            )

            thumb_path = (
                THUMB_FOLDER /
                image_name
            )

            image_urls.append({

                "original":
                    original_path.as_posix(),

                "thumbnail":
                    thumb_path.as_posix()

            })

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
                    row["採　　集　日"]
                ).strip(),

            "notes":
                str(
                    row["備考"]
                ).strip(),

            "images":
                image_urls,

            "site_id":
                str(
                    row[SITE_ID]
                ).strip()

        }

        # --------------------------------------------------------
        # 後方互換用
        #
        # 現在のindex.htmlが
        # specimen.image
        # を参照している場合にも対応。
        # --------------------------------------------------------

        if image_urls:

            data["image"] = image_urls[0][
                "original"
            ]

        else:

            data["image"] = ""

        specimen_data.append(
            data
        )

    with open(
        OUTPUT_FOLDER /
        "specimens.json",
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            specimen_data,
            file,
            ensure_ascii=False,
            indent=2
        )

    print_ok(
        "specimens.jsonを作成しました。"
    )

    # ------------------------------------------------------------
    # localities_public.json
    # ------------------------------------------------------------

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

        if not coordinates:

            continue

        average_latitude = (

            sum(
                coordinate[0]
                for coordinate in coordinates
            )
            /
            len(coordinates)

        )

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
    ) as file:

        json.dump(
            locality_data,
            file,
            ensure_ascii=False,
            indent=2
        )

    print_ok(
        "localities_public.jsonを作成しました。"
    )

    return (
        specimen_data,
        locality_data
    )


# ================================================================
# Git差分確認
# ================================================================

def show_git_status():

    print_title("【9】Gitの変更内容を確認しています")

    success, output = run_git_command(
        [
            "git",
            "status",
            "--short"
        ]
    )

    if not success:

        return False

    if output.strip() == "":

        print_info(
            "Gitに新しい変更はありません。"
        )

        return True

    print(
        "今回GitHubへ送信される変更："
    )

    print(
        output
    )

    return True


# ================================================================
# Git add
# ================================================================

def git_add():

    print_title("【10】Gitへ変更を登録しています")

    success, _ = run_git_command(
        [
            "git",
            "add",
            "."
        ]
    )

    if not success:

        print_error(
            "git addに失敗しました。"
        )

        return False

    print_ok(
        "Git add完了"
    )

    return True


# ================================================================
# Git commit
# ================================================================

def git_commit():

    print_title("【11】Gitへコミットしています")

    success, _ = run_git_command(
        [
            "git",
            "diff",
            "--cached",
            "--quiet"
        ]
    )

    # returncode 0 → 差分なし
    # returncode 1 → 差分あり

    if success:

        print_info(
            "コミットする変更はありません。"
        )

        return True

    commit_message = (
        "Update fossil database automatically"
    )

    success, _ = run_git_command(
        [
            "git",
            "commit",
            "-m",
            commit_message
        ]
    )

    if not success:

        print_error(
            "git commitに失敗しました。"
        )

        return False

    print_ok(
        "Git commit完了"
    )

    return True


# ================================================================
# Git push
# ================================================================

def git_push():

    print_title("【12】GitHubへpushしています")

    print(
        "  リモート：",
        GIT_REMOTE
    )

    print(
        "  ブランチ：",
        GIT_BRANCH
    )

    print()

    success, _ = run_git_command(
        [
            "git",
            "push",
            GIT_REMOTE,
            GIT_BRANCH
        ]
    )

    if not success:

        print_error(
            "GitHubへのpushに失敗しました。"
        )

        print()

        print(
            "GitHub認証が必要な場合は、"
            "ブラウザ認証を完了してください。"
        )

        return False

    print_ok(
        "GitHubへのpushが完了しました。"
    )

    return True


# ================================================================
# メイン処理
# ================================================================

def main():

    print()
    print("=" * 60)
    print(" 鵡川高校 化石デジタルライブラリー")
    print(" 自動更新プログラム")
    print("=" * 60)
    print()

    print(
        "このプログラムは以下を自動実行します。"
    )

    print(
        "  1. データチェック"
    )

    print(
        "  2. 公開JSON生成"
    )

    print(
        "  3. サムネイル生成"
    )

    print(
        "  4. Git commit"
    )

    print(
        "  5. GitHub push"
    )

    print()

    # ------------------------------------------------------------
    # Gitリポジトリ確認
    # ------------------------------------------------------------

    if not check_git_repository():

        input(
            "\nEnterキーを押して終了してください。"
        )

        return

    # ------------------------------------------------------------
    # Excel
    # ------------------------------------------------------------

    specimens, localities = load_excel()

    if specimens is None:

        input(
            "\nEnterキーを押して終了してください。"
        )

        return

    # ------------------------------------------------------------
    # データチェック
    # ------------------------------------------------------------

    all_errors = []

    errors = check_columns(
        specimens,
        localities
    )

    all_errors.extend(
        errors
    )

    if errors:

        print()

        print_error(
            "列名に問題があるため処理を中止します。"
        )

        input(
            "\nEnterキーを押して終了してください。"
        )

        return

    errors = check_specimen_ids(
        specimens
    )

    all_errors.extend(
        errors
    )

    errors = check_site_ids(
        specimens,
        localities
    )

    all_errors.extend(
        errors
    )

    errors = check_images(
        specimens
    )

    all_errors.extend(
        errors
    )

    coordinate_errors, site_coordinates = (
        check_coordinates(
            localities
        )
    )

    all_errors.extend(
        coordinate_errors
    )

    # ------------------------------------------------------------
    # エラーがあれば公開データを作らない
    # ------------------------------------------------------------

    if all_errors:

        print_title(
            "★★★ データチェックエラー ★★★"
        )

        print(
            "エラー件数：",
            len(all_errors),
            "件"
        )

        print()

        print(
            "安全のため、"
            "JSON生成・GitHubへのpushを中止します。"
        )

        print()

        input(
            "Enterキーを押して終了してください。"
        )

        return

    # ------------------------------------------------------------
    # サムネイル
    # ------------------------------------------------------------

    if not create_thumbnails():

        print()

        print_error(
            "サムネイル作成に失敗したため、"
            "GitHubへのpushを中止します。"
        )

        input(
            "\nEnterキーを押して終了してください。"
        )

        return

    # ------------------------------------------------------------
    # JSON
    # ------------------------------------------------------------

    specimen_data, locality_data = (
        create_public_json(
            specimens,
            localities,
            site_coordinates
        )
    )

    # ------------------------------------------------------------
    # Git status
    # ------------------------------------------------------------

    if not show_git_status():

        input(
            "\nEnterキーを押して終了してください。"
        )

        return

    # ------------------------------------------------------------
    # Git add
    # ------------------------------------------------------------

    if not git_add():

        input(
            "\nEnterキーを押して終了してください。"
        )

        return

    # ------------------------------------------------------------
    # Git commit
    # ------------------------------------------------------------

    if not git_commit():

        input(
            "\nEnterキーを押して終了してください。"
        )

        return

    # ------------------------------------------------------------
    # Git push
    # ------------------------------------------------------------

    if not git_push():

        input(
            "\nEnterキーを押して終了してください。"
        )

        return

    # ------------------------------------------------------------
    # 完了
    # ------------------------------------------------------------

    print()
    print("=" * 60)
    print(" ★★★ 自動更新完了 ★★★")
    print("=" * 60)
    print()

    print(
        "標本数：",
        len(specimen_data),
        "件"
    )

    print(
        "公開地点数：",
        len(locality_data),
        "件"
    )

    print()

    print(
        "完了した処理："
    )

    print(
        "  ✓ Excelデータチェック"
    )

    print(
        "  ✓ 画像チェック"
    )

    print(
        "  ✓ サムネイル生成"
    )

    print(
        "  ✓ specimens.json更新"
    )

    print(
        "  ✓ localities_public.json更新"
    )

    print(
        "  ✓ Git commit"
    )

    print(
        "  ✓ GitHub push"
    )

    print()

    print(
        "GitHub Pagesは自動的に更新されます。"
    )

    print()

    input(
        "Enterキーを押して終了してください。"
    )


# ================================================================
# 実行
# ================================================================

if __name__ == "__main__":

    main()
```

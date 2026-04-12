#!/usr/bin/env python3
"""
WWOOF Japan ホスト検索スクレイパー
指定した月に受け入れ可能なホストを検索してCSV/JSONに出力します。

使い方:
    python wwoof_japan_scraper.py --username <メールアドレス> --password <パスワード> --month 5

依存:
    pip install requests beautifulsoup4 lxml
"""

import argparse
import csv
import json
import re
import sys
import time
from dataclasses import asdict, dataclass, field
from typing import Optional

import requests
from bs4 import BeautifulSoup

BASE_URL = "https://www.wwoofjapan.com/home/index.php"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
    "Referer": "https://www.wwoofjapan.com/",
}

MONTH_NAMES = {
    1: "1月", 2: "2月", 3: "3月", 4: "4月",
    5: "5月", 6: "6月", 7: "7月", 8: "8月",
    9: "9月", 10: "10月", 11: "11月", 12: "12月",
}


@dataclass
class Host:
    id: str = ""
    name: str = ""
    prefecture: str = ""
    city: str = ""
    farm_type: str = ""
    available_months: list = field(default_factory=list)
    max_woofers: str = ""
    min_stay_days: str = ""
    max_stay_days: str = ""
    languages: str = ""
    description: str = ""
    profile_url: str = ""


class WwoofJapanScraper:
    def __init__(self, username: str, password: str, lang: str = "jp"):
        self.username = username
        self.password = password
        self.lang = lang
        self.session = requests.Session()
        self.session.headers.update(HEADERS)

    # ------------------------------------------------------------------
    # ログイン
    # ------------------------------------------------------------------

    def login(self) -> bool:
        """WWOOFジャパンにログインします。成功したらTrueを返します。"""
        print("[*] ログインページを取得中...")
        login_url = (
            f"{BASE_URL}?option=com_comprofiler&task=login&lang={self.lang}"
        )
        resp = self.session.get(login_url, timeout=30)
        resp.raise_for_status()

        soup = BeautifulSoup(resp.text, "lxml")

        # CSRF トークン（Joomla の com_user / com_comprofiler 両方に対応）
        token = self._extract_token(soup)
        if not token:
            print("[!] CSRFトークンが見つかりません。ページ構造が変わっている可能性があります。")

        post_data: dict = {
            "option": "com_comprofiler",
            "task": "login",
            "lang": self.lang,
            "username": self.username,
            "passwd": self.password,
            "return": "",
            "message": "0",
        }
        if token:
            post_data[token] = "1"

        # com_user ログインフォームも試みる
        joomla_token = self._extract_joomla_token(soup)
        if joomla_token:
            post_data[joomla_token] = "1"

        print("[*] ログイン中...")
        resp = self.session.post(
            BASE_URL, data=post_data, timeout=30, allow_redirects=True
        )
        resp.raise_for_status()

        if self._is_logged_in(resp.text):
            print("[+] ログイン成功")
            return True

        # フォールバック: Joomla 標準ログイン
        return self._login_joomla_standard()

    def _login_joomla_standard(self) -> bool:
        """Joomla 標準ログインフォームでリトライ。"""
        login_url = f"{BASE_URL}?option=com_users&view=login&lang={self.lang}"
        resp = self.session.get(login_url, timeout=30)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "lxml")

        form = soup.find("form", id=re.compile(r"login", re.I))
        if not form:
            form = soup.find("form")

        post_data: dict = {
            "option": "com_users",
            "task": "user.login",
            "username": self.username,
            "password": self.password,
            "return": "",
        }

        token = self._extract_joomla_token(soup)
        if token:
            post_data[token] = "1"

        action = form.get("action", BASE_URL) if form else BASE_URL

        resp = self.session.post(
            action, data=post_data, timeout=30, allow_redirects=True
        )
        resp.raise_for_status()

        if self._is_logged_in(resp.text):
            print("[+] ログイン成功 (Joomla標準)")
            return True

        print("[!] ログイン失敗。ユーザー名・パスワードをご確認ください。")
        return False

    def _extract_token(self, soup: BeautifulSoup) -> Optional[str]:
        """Community Builder の hidden トークンフィールド名を返す。"""
        for inp in soup.find_all("input", {"type": "hidden"}):
            name = inp.get("name", "")
            value = inp.get("value", "")
            # CB トークンは 32文字の英数字が多い
            if len(name) == 32 and re.match(r"^[a-f0-9]+$", name):
                return name
        return None

    def _extract_joomla_token(self, soup: BeautifulSoup) -> Optional[str]:
        """Joomla CSRF トークン（value="1" の hidden input）を返す。"""
        for inp in soup.find_all("input", {"type": "hidden", "value": "1"}):
            name = inp.get("name", "")
            if len(name) == 32:
                return name
        return None

    def _is_logged_in(self, html: str) -> bool:
        """ログイン済みかどうかをHTMLから判定。"""
        return (
            "ログアウト" in html
            or "logout" in html.lower()
            or "log out" in html.lower()
            or "マイページ" in html
            or "task=logout" in html
        )

    # ------------------------------------------------------------------
    # ホスト一覧取得
    # ------------------------------------------------------------------

    def fetch_host_list(self, month: int) -> list[Host]:
        """指定した月に受け入れ可能なホスト一覧を返します。"""
        print(f"\n[*] {MONTH_NAMES[month]}受け入れ可能なホストを検索中...")

        all_hosts: list[Host] = []
        page = 0
        limit = 20  # 1ページあたりの件数

        while True:
            hosts, has_next = self._fetch_page(month, page, limit)
            all_hosts.extend(hosts)
            print(f"    ページ {page + 1}: {len(hosts)} 件取得")

            if not has_next or len(hosts) == 0:
                break

            page += 1
            time.sleep(1.5)  # サーバー負荷軽減

        print(f"[+] 合計 {len(all_hosts)} 件のホストを取得しました")
        return all_hosts

    def _fetch_page(self, month: int, page: int, limit: int) -> tuple[list[Host], bool]:
        """1ページ分のホスト一覧を取得。(hosts, has_next_page) を返す。"""
        params = {
            "option": "com_comprofiler",
            "view": "userslist",
            "searchmode": "1",
            "listid": "4",
            "Itemid": "2356",
            "lang": self.lang,
            "limitstart": str(page * limit),
            "limit": str(limit),
        }

        # 月フィールドは Community Builder のカスタムフィールド名
        # WWOOF Japan のフィールド名（調査結果に基づく）
        month_field_candidates = [
            f"search[cb_available_months][]",
            f"search[cb_month][]",
            f"search[cb_accept_month][]",
            f"cbsearch[available_months][]",
        ]

        # まず検索フォームを取得してフィールド名を特定
        resp = self.session.get(BASE_URL, params=params, timeout=30)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "lxml")

        month_field = self._detect_month_field(soup)
        post_params = dict(params)
        post_params["searchmode"] = "1"

        if month_field:
            post_params[month_field] = str(month)
            resp = self.session.post(BASE_URL, data=post_params, timeout=30)
            resp.raise_for_status()
            soup = BeautifulSoup(resp.text, "lxml")
        # フィールドが見つからない場合はGETの結果をそのまま使い、後でフィルタ

        hosts = self._parse_host_list(soup, month)
        has_next = self._has_next_page(soup, page, limit)
        return hosts, has_next

    def _detect_month_field(self, soup: BeautifulSoup) -> Optional[str]:
        """検索フォームから月フィールドの名前を検出。"""
        for inp in soup.find_all(["input", "select"]):
            name = inp.get("name", "")
            if any(kw in name.lower() for kw in ["month", "月", "available", "accept"]):
                return name
        return None

    def _parse_host_list(self, soup: BeautifulSoup, target_month: int) -> list[Host]:
        """ホスト一覧HTMLからHostオブジェクトのリストを生成。"""
        hosts: list[Host] = []

        # Community Builder のユーザーリスト構造
        # 各ホストは <div class="cb_result_row"> や <tr> に入ることが多い
        rows = (
            soup.find_all("div", class_=re.compile(r"cb_result|cbUserListRow|userlist_row", re.I))
            or soup.find_all("tr", class_=re.compile(r"row|host", re.I))
        )

        if not rows:
            # テーブル形式の場合
            table = soup.find("table", class_=re.compile(r"userlist|hostlist|cb_list", re.I))
            if table:
                rows = table.find_all("tr")[1:]  # ヘッダー行をスキップ

        for row in rows:
            host = self._parse_host_row(row)
            if host and host.id:
                # 月フィルタ（サーバーサイドフィルタが効かなかった場合）
                if not host.available_months or target_month in host.available_months:
                    hosts.append(host)

        # リンクからホストページを個別取得（詳細が一覧に含まれない場合）
        if not hosts:
            hosts = self._parse_host_links(soup, target_month)

        return hosts

    def _parse_host_row(self, row) -> Optional[Host]:
        """1行からHostオブジェクトを生成。"""
        host = Host()

        # プロフィールURL & ID
        link = row.find("a", href=re.compile(r"userProfile|user=", re.I))
        if link:
            href = link.get("href", "")
            host.profile_url = href if href.startswith("http") else f"https://www.wwoofjapan.com{href}"
            m = re.search(r"user=(\d+)", href)
            if m:
                host.id = m.group(1)

        # 名前
        name_el = row.find(class_=re.compile(r"name|title|username", re.I))
        if name_el:
            host.name = name_el.get_text(strip=True)
        elif link:
            host.name = link.get_text(strip=True)

        # 都道府県・地域
        pref_el = row.find(class_=re.compile(r"pref|prefecture|location|region", re.I))
        if pref_el:
            host.prefecture = pref_el.get_text(strip=True)

        # テキスト全体から都道府県を抽出
        text = row.get_text(" ", strip=True)
        pref_match = re.search(
            r"(北海道|青森|岩手|宮城|秋田|山形|福島|茨城|栃木|群馬|埼玉|千葉|東京|神奈川"
            r"|新潟|富山|石川|福井|山梨|長野|岐阜|静岡|愛知|三重|滋賀|京都|大阪|兵庫"
            r"|奈良|和歌山|鳥取|島根|岡山|広島|山口|徳島|香川|愛媛|高知|福岡|佐賀"
            r"|長崎|熊本|大分|宮崎|鹿児島|沖縄)[都道府県]?",
            text,
        )
        if pref_match and not host.prefecture:
            host.prefecture = pref_match.group(0)

        return host if host.id else None

    def _parse_host_links(self, soup: BeautifulSoup, target_month: int) -> list[Host]:
        """個別リンクからホストプロフィールを巡回して取得。"""
        hosts: list[Host] = []
        links = soup.find_all("a", href=re.compile(r"userProfile|task=userProfile", re.I))

        for link in links:
            href = link.get("href", "")
            url = href if href.startswith("http") else f"https://www.wwoofjapan.com{href}"
            m = re.search(r"user=(\d+)", href)
            if not m:
                continue
            user_id = m.group(1)

            host = self._fetch_host_profile(user_id, url, target_month)
            if host:
                hosts.append(host)
                time.sleep(1.0)

        return hosts

    # ------------------------------------------------------------------
    # ホスト詳細取得
    # ------------------------------------------------------------------

    def _fetch_host_profile(
        self, user_id: str, url: str, target_month: int
    ) -> Optional[Host]:
        """個別ホストプロフィールページを取得してHostを返す。"""
        try:
            resp = self.session.get(url, timeout=30)
            resp.raise_for_status()
        except requests.RequestException as e:
            print(f"    [!] プロフィール取得失敗 (user={user_id}): {e}")
            return None

        soup = BeautifulSoup(resp.text, "lxml")
        host = Host(id=user_id, profile_url=url)

        # ホスト名
        h1 = soup.find("h1") or soup.find("h2")
        if h1:
            host.name = h1.get_text(strip=True)

        text = soup.get_text(" ", strip=True)

        # 都道府県
        pref_match = re.search(
            r"(北海道|青森|岩手|宮城|秋田|山形|福島|茨城|栃木|群馬|埼玉|千葉|東京|神奈川"
            r"|新潟|富山|石川|福井|山梨|長野|岐阜|静岡|愛知|三重|滋賀|京都|大阪|兵庫"
            r"|奈良|和歌山|鳥取|島根|岡山|広島|山口|徳島|香川|愛媛|高知|福岡|佐賀"
            r"|長崎|熊本|大分|宮崎|鹿児島|沖縄)[都道府県]?",
            text,
        )
        if pref_match:
            host.prefecture = pref_match.group(0)

        # 受け入れ可能月（チェックボックスや "5月" のようなテキスト）
        host.available_months = self._extract_months(soup, text)

        # 農場タイプ
        farm_match = re.search(
            r"(農業|有機農業|野菜|果樹|米|酪農|養蜂|ハーブ|畜産|林業|漁業|観光農園|ワイナリー"
            r"|カフェ|宿泊|アウトドア|自給自足)[^\s　]{0,20}",
            text,
        )
        if farm_match:
            host.farm_type = farm_match.group(0)

        # 最大受け入れ人数
        max_match = re.search(r"最大[^\d]*(\d+)\s*人", text)
        if max_match:
            host.max_woofers = max_match.group(1)

        # 最低滞在日数
        min_stay = re.search(r"最低[^\d]*(\d+)\s*日", text)
        if min_stay:
            host.min_stay_days = min_stay.group(1)

        # 目標月に受け入れ可能かフィルタ
        if host.available_months and target_month not in host.available_months:
            return None

        return host

    def _extract_months(self, soup: BeautifulSoup, text: str) -> list[int]:
        """HTMLとテキストから受け入れ可能月を抽出。"""
        months: list[int] = []

        # チェックボックス/選択済みフィールドから
        for inp in soup.find_all("input", {"type": "checkbox", "checked": True}):
            val = inp.get("value", "")
            if val.isdigit() and 1 <= int(val) <= 12:
                months.append(int(val))

        # テキストパターン: "5月・6月・7月" や "5,6,7月"
        if not months:
            raw = re.findall(r"(\d{1,2})[月,・/]", text)
            for v in raw:
                if 1 <= int(v) <= 12:
                    months.append(int(v))

        return sorted(set(months))

    # ------------------------------------------------------------------
    # 出力
    # ------------------------------------------------------------------

    def save_csv(self, hosts: list[Host], filepath: str) -> None:
        """ホスト一覧をCSVに保存。"""
        if not hosts:
            print("[!] 保存するデータがありません")
            return

        fieldnames = [
            "id", "name", "prefecture", "city", "farm_type",
            "available_months", "max_woofers", "min_stay_days",
            "max_stay_days", "languages", "description", "profile_url",
        ]

        with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for host in hosts:
                row = asdict(host)
                row["available_months"] = "・".join(
                    f"{m}月" for m in row["available_months"]
                )
                writer.writerow(row)

        print(f"[+] CSV保存: {filepath}")

    def save_json(self, hosts: list[Host], filepath: str) -> None:
        """ホスト一覧をJSONに保存。"""
        data = [asdict(h) for h in hosts]
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"[+] JSON保存: {filepath}")

    def print_summary(self, hosts: list[Host], month: int) -> None:
        """コンソールに概要を表示。"""
        print(f"\n{'='*60}")
        print(f"  {MONTH_NAMES[month]}に申し込み可能なホスト一覧")
        print(f"{'='*60}")
        if not hosts:
            print("  該当するホストが見つかりませんでした。")
            return

        # 都道府県ごとにグループ化
        by_pref: dict[str, list[Host]] = {}
        for h in hosts:
            pref = h.prefecture or "不明"
            by_pref.setdefault(pref, []).append(h)

        for pref, pref_hosts in sorted(by_pref.items()):
            print(f"\n【{pref}】")
            for h in pref_hosts:
                months_str = (
                    "・".join(f"{m}月" for m in h.available_months)
                    if h.available_months
                    else "不明"
                )
                print(f"  ▸ {h.name or '(名前なし)'}")
                print(f"      受け入れ月 : {months_str}")
                if h.farm_type:
                    print(f"      農場タイプ : {h.farm_type}")
                if h.max_woofers:
                    print(f"      最大人数   : {h.max_woofers}人")
                print(f"      URL        : {h.profile_url}")

        print(f"\n合計: {len(hosts)} 件")


# ------------------------------------------------------------------
# メイン
# ------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="WWOOF Japan ホスト検索スクレイパー"
    )
    parser.add_argument("--username", "-u", required=True, help="ログインメールアドレス")
    parser.add_argument("--password", "-p", required=True, help="ログインパスワード")
    parser.add_argument(
        "--month", "-m", type=int, default=5,
        help="受け入れ月 (1-12, デフォルト: 5)"
    )
    parser.add_argument("--output-csv", default="wwoof_hosts_may.csv", help="CSV出力ファイル名")
    parser.add_argument("--output-json", default="wwoof_hosts_may.json", help="JSON出力ファイル名")
    parser.add_argument("--no-save", action="store_true", help="ファイルに保存しない")
    args = parser.parse_args()

    if not 1 <= args.month <= 12:
        print("[!] --month は 1〜12 の整数を指定してください")
        sys.exit(1)

    scraper = WwoofJapanScraper(args.username, args.password)

    if not scraper.login():
        print("[!] ログインに失敗しました。終了します。")
        sys.exit(1)

    hosts = scraper.fetch_host_list(month=args.month)

    scraper.print_summary(hosts, args.month)

    if not args.no_save and hosts:
        scraper.save_csv(hosts, args.output_csv)
        scraper.save_json(hosts, args.output_json)


if __name__ == "__main__":
    main()

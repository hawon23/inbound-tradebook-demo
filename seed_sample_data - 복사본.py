"""
가상 샘플 데이터 시딩 스크립트
------------------------------------
app.py와 같은 폴더에서 아래처럼 한 번만 실행하면, tradebook.db에
가상 계좌/가상 ETF로 구성된 샘플 거래 6건 + 휴일 2건이 들어갑니다.
실제 거래 데이터는 전혀 사용하지 않았습니다 (계좌번호/펀드명/티커 모두 가상).

실행 방법:
    python seed_sample_data.py

이후 `streamlit run app.py`로 앱을 켜면, "다가오는 Reset" 페이지 등에서
바로 동작하는 화면을 확인/시연할 수 있습니다.
"""

import sqlite3
import datetime

DB_PATH = "tradebook.db"


def main():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    reset_cols_sql = ", ".join(f"reset_{i} TEXT" for i in range(1, 15))
    cur.executescript(
        f"""
        CREATE TABLE IF NOT EXISTS trades (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kis_ref_no TEXT UNIQUE,
            book TEXT NOT NULL DEFAULT 'End Client',
            status TEXT NOT NULL DEFAULT 'open',
            note TEXT,
            next_valuation_date TEXT,
            counterparty TEXT,
            trs_position TEXT,
            trade_date TEXT,
            effective_date TEXT,
            final_valuation_date TEXT,
            final_settlement_date TEXT,
            underlying_code TEXT,
            underlying_description TEXT,
            number_of_units REAL,
            initial_price REAL,
            floating_rate_index TEXT,
            spread REAL,
            equity_notional_amount REAL,
            net_number_of_units REAL,
            {reset_cols_sql},
            partner_trade_id INTEGER,
            rollover_from_id INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS holidays (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT UNIQUE,
            name TEXT
        );
        CREATE TABLE IF NOT EXISTS unwinds (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            trs_code TEXT,
            unwind_date TEXT,
            unwind_size REAL,
            unwind_price REAL,
            sales_tax TEXT NOT NULL DEFAULT 'N',
            settlement_days INTEGER NOT NULL DEFAULT 4,
            interest_rate_override REAL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        """
    )

    today = datetime.date.today()

    def d(offset):
        return (today + datetime.timedelta(days=offset)).isoformat()

    sample_trades = [
        ("TRS26SP00101", "End Client", "ETF A 신규", "A자산운용", "LONG", d(-60), d(-59),
         "000001 XX Equity", "샘플종목 A", 10000, 20000.0, "KWCDC Curncy", -0.005, 200000000,
         [d(-30), d(3), None, None]),
        ("TRS26SP00102", "BTB", "ETF A 짝꿍(GS)", "GS", "LONG", d(-60), d(-59),
         "000001 XX Equity", "샘플종목 A", 10000, 20000.0, "KWCDC Curncy", 0.010, 200000000,
         [d(-30), d(3), None, None]),
        ("TRS26SP00201", "End Client", "ETF B 신규", "A자산운용", "LONG", d(-40), d(-39),
         "000002 XX Equity", "샘플종목 B", 5000, 18000.0, "KWCDC Curncy", -0.004, 90000000,
         [d(8), None, None, None]),
        ("TRS26SP00202", "BTB", "ETF B 짝꿍(Nomura)", "Nomura", "LONG", d(-40), d(-39),
         "000002 XX Equity", "샘플종목 B", 5000, 18000.0, "KWCDC Curncy", 0.008, 90000000,
         [d(8), None, None, None]),
        ("TRS26SP00301", "End Client", "ETF C 신규(지난 리셋 포함)", "A자산운용", "LONG", d(-90), d(-89),
         "000003 XX Equity", "샘플종목 C", 8000, 22000.0, "KWCDC Curncy", -0.005, 176000000,
         [d(-45), d(-2), d(13), None]),
        ("TRS26SP00302", "BTB", "ETF C 짝꿍(GS)", "GS", "LONG", d(-90), d(-89),
         "000003 XX Equity", "샘플종목 C", 8000, 22000.0, "KWCDC Curncy", 0.010, 176000000,
         [d(-45), d(-2), d(13), None]),
    ]

    cols = [
        "kis_ref_no", "book", "note", "counterparty", "trs_position", "trade_date",
        "effective_date", "underlying_code", "underlying_description", "number_of_units",
        "initial_price", "floating_rate_index", "spread", "equity_notional_amount",
        "net_number_of_units", "reset_1", "reset_2", "reset_3", "reset_4",
    ]
    for row in sample_trades:
        (ref, book, note, cp, pos, tdate, edate, ucode, udesc, units, price, fri,
         spread, notional, resets) = row
        vals = [ref, book, note, cp, pos, tdate, edate, ucode, udesc, units, price,
                fri, spread, notional, units] + resets
        placeholders = ",".join(["?"] * len(cols))
        cur.execute(
            f"INSERT OR IGNORE INTO trades ({','.join(cols)}) VALUES ({placeholders})",
            vals,
        )

    for hd, name in [(d(5), "샘플 공휴일"), (d(20), "샘플 공휴일2")]:
        cur.execute("INSERT OR IGNORE INTO holidays (date, name) VALUES (?,?)", (hd, name))

    conn.commit()
    cur.execute("SELECT count(*) FROM trades")
    print(f"완료: trades {cur.fetchone()[0]}건 저장됨 (tradebook.db)")
    conn.close()


if __name__ == "__main__":
    main()

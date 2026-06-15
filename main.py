from pathlib import Path

from app.core.paths import get_base_dir
from app.infrastructure.excel_runner import ExcelRunner

TEST_QUERY = """
SELECT
    pf.PrefecturalID
    , pf.PrefecturalName
FROM 
    Prefecturals pf
WHERE 
    pf.PrefecturalID IN (?, ?, ?)
;
"""

def main():
    runner = ExcelRunner()

    runner.execute(
        output_path=get_base_dir() / "temp" / "test.json", 
        query=TEST_QUERY, 
        params=[28, 29, 30], 
        timeout=30
    )

if __name__ == "__main__":
    main()

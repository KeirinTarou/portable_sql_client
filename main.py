from pathlib import Path

from app.core.paths import get_base_dir
from app.infrastructure.excel_runner import ExcelRunner
from app.infrastructure.json_loader import JSONLoader

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

    test_json_path = get_base_dir() / "temp" / "test.json"
    runner.execute(
        output_path=test_json_path, 
        query=TEST_QUERY, 
        params=[28, 29, 30], 
        timeout=30
    )
    loader = JSONLoader()
    result = loader.load(test_json_path)

    print(result.columns)
    print(result.rows)

if __name__ == "__main__":
    main()

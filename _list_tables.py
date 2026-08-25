import store

# 查看数据库表结构
tables = store.rows("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
print("数据库表:")
for t in tables:
    print(f"  {t['name']}")

# 查看offerings表
try:
    offerings = store.rows("SELECT * FROM offerings ORDER BY id")
    print(f"\nofferings: {len(offerings)} rows")
    if offerings:
        print(f"字段: {list(offerings[0].keys())}")
        for o in offerings[:5]:
            print(f"  {o}")
except:
    print("无offerings表")

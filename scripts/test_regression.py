import base64
import importlib
import os
import tempfile


def _urlsafe_b64_json(text: str) -> str:
    return base64.urlsafe_b64encode(text.encode("utf-8")).decode("ascii").rstrip("=")


def main():
    mod = importlib.import_module("backend_server")
    tmpdir = tempfile.mkdtemp(prefix="baoguan-regression-")
    generated_dir = os.path.join(tmpdir, "generated")
    os.makedirs(generated_dir, exist_ok=True)

    mod.CACHE_DB_PATH = os.path.join(generated_dir, "request_cache.sqlite3")
    mod.OUTPUT_DIR = generated_dir
    mod._TOKEN_LOCKS.clear()
    mod._init_cache_db()

    calls = {"count": 0}

    def fake_fill_template(rows):
        calls["count"] += 1
        filename = f"dummy-{calls['count']}.xlsx"
        with open(os.path.join(generated_dir, filename), "wb") as fh:
            fh.write(("rows=%d" % len(rows)).encode("utf-8"))
        return filename

    mod.fill_template = fake_fill_template

    client = mod.app.test_client()
    rows = [
        {
            "合同号码": "PK-TEST-001",
            "境外收货人": "Demo Receiver",
            "贸易国": "美国",
            "运抵国": "美国",
            "成交方式": "FOB",
            "商品编号": "8507600090",
            "品牌": "PKCELL",
            "规格型号": "IFR14500-800-3.2V",
            "数量": "10",
            "单位": "PCS",
            "单价": "1.23",
            "金额": "12.30",
            "币制": "USD",
            "货源地": "深圳",
        }
    ]

    resp = client.post(
        "/cache?trace_id=BG-TEST-001",
        json={"rows": rows, "meta": {"scriptVersion": "test"}},
    )
    assert resp.status_code == 200, resp.get_data(as_text=True)
    payload = resp.get_json()
    assert payload["traceId"] == "BG-TEST-001"
    assert "trace_id=BG-TEST-001" in payload["url"]
    token = payload["token"]

    resp1 = client.get(f"/generate?t={token}&trace_id=BG-TEST-001")
    assert resp1.status_code == 200, resp1.get_data(as_text=True)
    assert calls["count"] == 1
    resp1.close()

    resp2 = client.get(f"/generate?t={token}&trace_id=BG-TEST-001")
    assert resp2.status_code == 200, resp2.get_data(as_text=True)
    assert calls["count"] == 1
    resp2.close()

    legacy_data = _urlsafe_b64_json('{"rows":[{"合同号码":"PK-LEGACY-001"}]}')
    resp3 = client.get(f"/generate?d={legacy_data}&trace_id=BG-LEGACY-001")
    assert resp3.status_code == 200, resp3.get_data(as_text=True)
    assert calls["count"] == 2
    resp3.close()

    print("regression ok")


if __name__ == "__main__":
    main()

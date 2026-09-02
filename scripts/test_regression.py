import base64
import importlib
import os
import tempfile


def _urlsafe_b64_json(text: str) -> str:
    return base64.urlsafe_b64encode(text.encode("utf-8")).decode("ascii").rstrip("=")


def main():
    mod = importlib.import_module("backend_server")
    fixed_cr_capacity_mah = {
        "CR1025": 30,
        "CR1220": 40,
        "CR1225": 50,
        "CR1616": 50,
        "CR1620": 70,
        "CR1632": 120,
        "CR2016": 76,
        "CR2025": 150,
        "CR2032": 210,
        "CR2320": 130,
        "CR2325": 210,
        "CR2330": 260,
        "CR2430": 270,
        "CR2450": 600,
        "CR2477": 1000,
        "CR3032": 580,
        "CR2": 850,
        "CR123A": 1500,
    }
    for model, expected_capacity in fixed_cr_capacity_mah.items():
        parsed = mod._parse_spec(f"PKCELL-{model}")
        assert parsed["model"] == model, (model, parsed)
        assert parsed["capacity_value"] == float(expected_capacity), (model, parsed)
        assert parsed["capacity_unit"] == "mah", (model, parsed)
        assert parsed["capacity_mah"] == float(expected_capacity), (model, parsed)
        assert parsed["voltage_v"] == 3.0, (model, parsed)

    for spec, expected_capacity in (
        ("PKCELL-CR2032-220mAh-3V", 220.0),
        ("PKCELL-CR2477-900mAh-3V", 900.0),
    ):
        parsed = mod._parse_spec(spec)
        assert parsed["capacity_mah"] == expected_capacity, parsed

    for spec in (
        "PKCELL-HPC1520-带焊片",
        "PKCELL-ER34615-19000-3.6V+HPC1520-T",
        "PKCELL-HPC1520-120mAh-3.6V",
    ):
        parsed = mod._parse_spec(spec)
        assert parsed["model"] == "HPC152", parsed
        assert parsed["capacity_mah"] == 90.0, parsed
        assert parsed["voltage_v"] == 4.0, parsed

    base_row = {
        "商品编号": "8506500090",
        "品牌": "PKCELL",
    }
    cr2_parts = mod.build_product_name({**base_row, "规格型号": "PKCELL-CR2"}).split("|")
    assert cr2_parts[3] == "圆柱形", cr2_parts
    assert cr2_parts[4] == "二氧化锰+铁+锂", cr2_parts
    button_parts = mod.build_product_name({**base_row, "规格型号": "PKCELL-CR1220"}).split("|")
    assert button_parts[3] == "纽扣形", button_parts
    assert mod._normalize_domestic_origin("常州其他 32049") == "常州其他 32049"

    lithium_ion_parts = mod.build_product_name(
        {
            "商品编号": "8507600099",
            "品牌": "PKCELL",
            "规格型号": "PKCELL-LP401230-105-3.7V-PT",
            "数量": "7000",
            "净重": "39.2",
        },
        pack_qty="7000",
        pack_net="39.2",
    ).split("|")
    assert lithium_ion_parts[4] == "锂离子", lithium_ion_parts
    assert lithium_ion_parts[6:10] == ["LP401230", "105mAh", "不含汞", "3.7V"], lithium_ion_parts
    assert lithium_ion_parts[10] == "比能量：69.38WH/KG", lithium_ion_parts

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
    health_resp = client.get("/health?trace_id=BG-HEALTH-TEST")
    assert health_resp.status_code == 200
    assert health_resp.get_json() == {"status": "ok", "storageBackend": "sqlite"}
    rows = [
        {
            "合同号码": "PK-TEST-001",
            "境外收货人": "Demo Receiver",
            "贸易国": "美国",
            "运抵国": "美国",
            "成交方式": "FOB",
            "商品编号": "8507600099",
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

    batch_items = []
    for index in range(3):
        batch_rows = [dict(rows[0], **{"合同号码": f"PK-BATCH-{index + 1:03d}"})]
        batch_items.append(
            {
                "rows": batch_rows,
                "meta": {
                    "traceId": f"BG-BATCH-{index + 1:03d}",
                    "scriptVersion": "batch-test",
                },
            }
        )
    batch_items.insert(
        1,
        {
            "rows": [],
            "meta": {"traceId": "BG-BATCH-INVALID", "scriptVersion": "batch-test"},
        },
    )

    batch_resp = client.post(
        "/generate?cache=1&batch=1&trace_id=BG-BATCH",
        json={"items": batch_items, "meta": {"scriptVersion": "batch-test"}},
    )
    assert batch_resp.status_code == 200, batch_resp.get_data(as_text=True)
    batch_payload = batch_resp.get_json()
    assert batch_payload["successCount"] == 3, batch_payload
    assert batch_payload["failureCount"] == 1, batch_payload
    assert [item["traceId"] for item in batch_payload["items"]] == [
        "BG-BATCH-001",
        "BG-BATCH-INVALID",
        "BG-BATCH-002",
        "BG-BATCH-003",
    ]
    assert batch_payload["items"][1]["ok"] is False, batch_payload
    assert batch_payload["items"][1]["code"] == "BG4003", batch_payload
    valid_batch_results = [item for item in batch_payload["items"] if item["ok"]]
    assert len({item["token"] for item in valid_batch_results}) == 3
    assert all("trace_id=" in item["url"] for item in valid_batch_results)

    empty_batch_resp = client.post("/generate?cache=1&batch=1", json={"items": []})
    assert empty_batch_resp.status_code == 400
    assert empty_batch_resp.get_json()["code"] == "BG4005"

    oversized_batch_resp = client.post(
        "/generate?cache=1&batch=1",
        json={"items": [{} for _ in range(mod.MAX_BATCH_CACHE_ITEMS + 1)]},
    )
    assert oversized_batch_resp.status_code == 400
    assert oversized_batch_resp.get_json()["code"] == "BG4005"

    options_resp = client.options(
        "/generate?cache=1&batch=1",
        headers={
            "Origin": "https://alidocs.dingtalk.com",
            "Access-Control-Request-Method": "POST",
            "Access-Control-Request-Headers": "content-type",
        },
    )
    assert options_resp.status_code == 204
    assert options_resp.headers["Access-Control-Max-Age"] == "600"

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

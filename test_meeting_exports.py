import unittest

from meeting_exports import _build_overview_svg, _ppt_overview_slide


def make_task_rows(level2_count=9, grandchildren=None):
    rows = [
        {"node_id": "N001", "canonical_name": "測試集團", "chart1_parent": "", "chart1_level": 0},
        {"node_id": "N002", "canonical_name": "測試控股有限公司", "chart1_parent": "N001", "chart1_level": 1},
    ]
    for idx in range(level2_count):
        node_id = f"N{idx + 3:03d}"
        rows.append(
            {
                "node_id": node_id,
                "canonical_name": f"同層子公司{idx + 1:02d}",
                "chart1_parent": "N002",
                "chart1_level": 2,
                "actual_controller_share": "100%",
            }
        )
        for child_idx in range((grandchildren or {}).get(idx + 1, 0)):
            rows.append(
                {
                    "node_id": f"N{idx + 3:03d}-{child_idx + 1}",
                    "canonical_name": f"下層公司{idx + 1:02d}-{child_idx + 1}",
                    "chart1_parent": node_id,
                    "chart1_level": 3,
                    "actual_controller_share": "100%",
                }
            )
    return rows


class MeetingOverviewTests(unittest.TestCase):
    def test_html_overview_draws_all_same_level_nodes_when_twelve_or_fewer(self):
        svg = _build_overview_svg("測試", make_task_rows(level2_count=9))

        for idx in range(1, 10):
            self.assertIn(f"同層子公司{idx:02d}", svg)
        self.assertNotIn("其他一級子公司", svg)
        self.assertNotIn("同層清單", svg)

    def test_ppt_overview_draws_all_same_level_nodes_when_twelve_or_fewer(self):
        slide_xml = _ppt_overview_slide("測試", make_task_rows(level2_count=9))

        for idx in range(1, 10):
            self.assertIn(f"同層子公司{idx:02d}", slide_xml)
        self.assertNotIn("其他一級子公司", slide_xml)
        self.assertNotIn("同層清單", slide_xml)

    def test_overview_uses_separate_peer_list_only_after_twelve_same_level_nodes(self):
        svg = _build_overview_svg("測試", make_task_rows(level2_count=14))

        for idx in range(1, 13):
            self.assertIn(f"同層子公司{idx:02d}", svg)
        self.assertIn("同層清單｜2 家", svg)
        self.assertNotIn("其他一級子公司", svg)

    def test_large_overview_switches_to_core_summary(self):
        svg = _build_overview_svg("測試", make_task_rows(level2_count=14, grandchildren={1: 4, 2: 4, 3: 4, 4: 4}))

        self.assertIn("核心摘要圖", svg)
        self.assertIn("完整層級請見清單", svg)
        for idx in range(1, 7):
            self.assertIn(f"同層子公司{idx:02d}", svg)
        self.assertNotIn("同層子公司07", svg)
        self.assertIn("另有 8 家同層分支未展開", svg)

    def test_large_ppt_overview_switches_to_core_summary(self):
        slide_xml = _ppt_overview_slide("測試", make_task_rows(level2_count=14, grandchildren={1: 4, 2: 4, 3: 4, 4: 4}))

        self.assertIn("核心摘要圖", slide_xml)
        self.assertIn("完整層級請見清單", slide_xml)
        for idx in range(1, 7):
            self.assertIn(f"同層子公司{idx:02d}", slide_xml)
        self.assertNotIn("同層子公司07", slide_xml)
        self.assertIn("另有 8 家同層分支未展開", slide_xml)


if __name__ == "__main__":
    unittest.main()

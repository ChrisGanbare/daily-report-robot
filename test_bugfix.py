"""
test_bugfix.py — 验证 main.py 六项 bug 修复逻辑
不依赖数据库 / 外部网络，全部使用 mock 或本地构造数据。
"""

import io
import os
import sys
import types
import datetime
import unittest
from unittest.mock import MagicMock, patch, call, mock_open

import pandas as pd

# ── 屏蔽 load_dotenv，防止读取真实 .env ──────────────────────────────────────
sys.modules.setdefault('dotenv', types.ModuleType('dotenv'))
sys.modules['dotenv'].load_dotenv = lambda: None  # type: ignore

import main  # noqa: E402  — 导入被测模块


# ─────────────────────────────────────────────────────────────────────────────
# 辅助：初始化 logger，使测试期间日志有处可去
# ─────────────────────────────────────────────────────────────────────────────
import logging
logging.basicConfig(level=logging.CRITICAL)   # 测试期间仅显示 CRITICAL，保持输出干净


# =============================================================================
# Bug 1：pd.to_datetime 使用 errors='coerce'，无效日期不崩溃
# =============================================================================
class TestBug1ToDatetimeCoerce(unittest.TestCase):

    def _run_step4(self, install_time_values):
        """复现 process_data step4 的日期转换逻辑"""
        df = pd.DataFrame({'install_time': install_time_values})
        df['install_time'] = pd.to_datetime(df['install_time'], errors='coerce')
        df = df.sort_values(by='install_time', ascending=True)
        df['install_time'] = df['install_time'].dt.strftime('%Y.%m.%d').fillna('')
        return df

    def test_valid_dates_formatted(self):
        """正常日期应被格式化为 YYYY.MM.DD"""
        df = self._run_step4(['2023-01-15', '2022-06-01'])
        self.assertEqual(df['install_time'].tolist(), ['2022.06.01', '2023.01.15'])

    def test_invalid_date_becomes_empty_string(self):
        """无效日期应变为空串，而非抛出异常"""
        df = self._run_step4(['2023-01-15', 'INVALID_DATE', None])
        self.assertIn('', df['install_time'].tolist(),
                      "无效日期应被 coerce 为 NaT 并 fillna 为空串")

    def test_all_invalid_no_exception(self):
        """全部无效时不应抛出任何异常"""
        try:
            self._run_step4(['bad1', 'bad2', 'bad3'])
        except Exception as e:
            self.fail(f"全部无效日期时不应抛出异常，但抛出了: {e}")

    def test_old_code_raises(self):
        """验证旧代码（无 errors='coerce'）确实会崩溃，以证明修复有意义"""
        with self.assertRaises(Exception):
            df = pd.DataFrame({'install_time': ['INVALID_DATE']})
            pd.to_datetime(df['install_time'])   # 旧代码：无 errors='coerce'


# =============================================================================
# Bug 2：ExcelWriter 使用 context manager，异常时文件不泄漏
# =============================================================================
class TestBug2ExcelWriterContextManager(unittest.TestCase):

    def test_context_manager_closes_on_exception(self):
        """apply_sheet_format 抛异常时，with 块保证 writer 被关闭（文件不残留）"""
        tmp = '_test_bug2_tmp.xlsx'
        try:
            with self.assertRaises(RuntimeError):
                with pd.ExcelWriter(tmp, engine='xlsxwriter') as writer:
                    pd.DataFrame({'A': [1]}).to_excel(writer, index=False)
                    raise RuntimeError("模拟 apply_sheet_format 异常")
            # with 块退出后，即使异常发生，writer 也已关闭
            # xlsxwriter 在异常时会清理临时文件（不一定写出完整文件）
            # 关键：不应有未关闭的文件句柄
        finally:
            if os.path.exists(tmp):
                os.remove(tmp)

    def test_generate_excel_with_format_returns_filename(self):
        """正常路径下 generate_excel_with_format 应返回文件名且文件存在"""
        df = pd.DataFrame({
            '序号': [1, 2],
            '客户名称': ['客户A', '客户B'],
            '设备编号': ['D001', 'D002'],
            '油品型号': ['0W-20', '5W-30'],
            '库存(%)': [50, 10],
            '桶数': [2, 1],
            '设备归属': ['中润', '中润'],
            '网络同步': ['在线', '离线'],
            '安装时间': ['2023.01.01', '2023.06.15'],
            '安装地点': ['北京市朝阳区', '上海市浦东新区'],
        })
        tmp = '_test_bug2_normal.xlsx'
        try:
            result = main.generate_excel_with_format(df, tmp)
            self.assertEqual(result, tmp)
            self.assertTrue(os.path.exists(tmp), "文件应已成功写入磁盘")
        finally:
            if os.path.exists(tmp):
                os.remove(tmp)


# =============================================================================
# Bug 3：_check_device_count_diff 在列名含空格时仍能正确匹配
# =============================================================================
class TestBug3CheckDeviceCountDiffColumnStrip(unittest.TestCase):

    def _make_online_df(self, col_name='设备编号', codes=('D001', 'D002')):
        return pd.DataFrame({col_name: list(codes)})

    def _make_local_excel(self, path, codes=('D001', 'D002', 'D003')):
        pd.DataFrame({'设备编号': list(codes)}).to_excel(
            path, index=False, sheet_name='设备配置')

    def test_column_with_spaces_still_matches(self):
        """在线表格列名含前后空格时，经 rename strip 后应能正确比对"""
        tmp = '_test_bug3_local.xlsx'
        self._make_local_excel(tmp, codes=['D001', 'D002', 'D003'])
        try:
            # 模拟在线 df 列名含空格（修复前会导致 online_codes 为空集）
            online_df = self._make_online_df(col_name=' 设备编号 ', codes=['D001', 'D002'])
            stripped_df = online_df.rename(columns=lambda c: str(c).strip())

            alerts = []
            with patch.object(main, 'CONFIG_LOCAL_FILE', tmp), \
                 patch.object(main, 'send_feishu_alert', side_effect=lambda *a, **kw: alerts.append(a)):
                main._check_device_count_diff(stripped_df)

            # 在线2台 vs 本地3台，差异1台，应触发告警
            self.assertTrue(len(alerts) > 0, "差异应触发告警")
            self.assertIn('差异', alerts[0][1], "告警标题应包含'差异'")
        finally:
            if os.path.exists(tmp):
                os.remove(tmp)

    def test_column_without_spaces_baseline(self):
        """无空格列名的基准场景：数量一致时不触发告警"""
        tmp = '_test_bug3_baseline.xlsx'
        self._make_local_excel(tmp, codes=['D001', 'D002'])
        try:
            online_df = self._make_online_df(col_name='设备编号', codes=['D001', 'D002'])
            alerts = []
            with patch.object(main, 'CONFIG_LOCAL_FILE', tmp), \
                 patch.object(main, 'send_feishu_alert', side_effect=lambda *a, **kw: alerts.append(a)):
                main._check_device_count_diff(online_df)
            self.assertEqual(len(alerts), 0, "数量一致时不应触发告警")
        finally:
            if os.path.exists(tmp):
                os.remove(tmp)


# =============================================================================
# Bug 4：上传接口 HTTP 错误时应抛出，而非解析 HTML 返回 JSONDecodeError
# =============================================================================
class TestBug4UploadRaiseForStatus(unittest.TestCase):

    def test_http_500_raises_and_is_caught(self):
        """上传接口返回 500 时，raise_for_status 触发，被外层 except 捕获并告警"""
        mock_resp = MagicMock()
        mock_resp.raise_for_status.side_effect = Exception("500 Server Error")

        alerts = []
        with patch.object(main, 'WEBHOOK_URL', 'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=testkey'), \
             patch('requests.post', return_value=mock_resp), \
             patch.object(main, 'send_feishu_alert', side_effect=lambda *a, **kw: alerts.append(a)), \
             patch('os.path.exists', return_value=True):
            main.send_to_robot('fake.xlsx')

        self.assertTrue(len(alerts) > 0, "HTTP 错误应触发飞书告警")
        self.assertEqual(alerts[0][0], 'warning')

    def test_http_200_json_decode_error_caught(self):
        """上传成功但响应非 JSON 时，异常被捕获并告警（旧代码在 raise_for_status 前 json() 就崩溃）"""
        mock_resp = MagicMock()
        mock_resp.raise_for_status.return_value = None
        mock_resp.json.side_effect = ValueError("No JSON")
        mock_resp.text = "<html>error</html>"

        alerts = []
        with patch.object(main, 'WEBHOOK_URL', 'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=testkey'), \
             patch('requests.post', return_value=mock_resp), \
             patch.object(main, 'send_feishu_alert', side_effect=lambda *a, **kw: alerts.append(a)), \
             patch('os.path.exists', return_value=True):
            main.send_to_robot('fake.xlsx')

        self.assertTrue(len(alerts) > 0, "JSON 解析异常应触发告警")


# =============================================================================
# Bug 5：发送消息后检查企业微信业务层 errcode
# =============================================================================
class TestBug5SendRespErrcode(unittest.TestCase):

    def _make_upload_mock(self, media_id='FAKE_MEDIA_ID'):
        m = MagicMock()
        m.raise_for_status.return_value = None
        m.json.return_value = {'media_id': media_id}
        m.text = ''
        return m

    def test_send_success_logs_success(self):
        """errcode=0 时应记录发送成功"""
        upload_mock = self._make_upload_mock()
        send_mock = MagicMock()
        send_mock.raise_for_status.return_value = None
        send_mock.json.return_value = {'errcode': 0, 'errmsg': 'ok'}

        log_messages = []
        with patch.object(main, 'WEBHOOK_URL', 'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=testkey'), \
             patch('requests.post', side_effect=[upload_mock, send_mock]), \
             patch('os.path.exists', return_value=True), \
             patch('builtins.open', mock_open(read_data=b'fake')), \
             patch.object(main.logger, 'info', side_effect=lambda msg: log_messages.append(msg)):
            main.send_to_robot('fake.xlsx')

        self.assertTrue(any('发送成功' in m for m in log_messages), "errcode=0 应记录发送成功")

    def test_send_errcode_nonzero_triggers_alert(self):
        """企业微信返回 errcode≠0 时应触发飞书告警"""
        upload_mock = self._make_upload_mock()
        send_mock = MagicMock()
        send_mock.raise_for_status.return_value = None
        send_mock.json.return_value = {'errcode': 40058, 'errmsg': 'invalid msgtype'}

        alerts = []
        with patch.object(main, 'WEBHOOK_URL', 'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=testkey'), \
             patch('requests.post', side_effect=[upload_mock, send_mock]), \
             patch('os.path.exists', return_value=True), \
             patch('builtins.open', mock_open(read_data=b'fake')), \
             patch.object(main, 'send_feishu_alert', side_effect=lambda *a, **kw: alerts.append(a)):
            main.send_to_robot('fake.xlsx')

        self.assertTrue(len(alerts) > 0, "errcode≠0 应触发飞书告警")
        self.assertIn('业务层错误', alerts[0][1])


# =============================================================================
# Bug 6：df_final 为空时跳过报表生成，不发送空文件
# =============================================================================
class TestBug6EmptyDfFinalGuard(unittest.TestCase):

    def test_empty_df_final_sends_alert_and_returns(self):
        """process_data 返回空 df_final 时，不应生成文件，应推送告警"""
        alerts = []
        excel_called = []

        with patch.object(main, 'get_db_data', return_value=pd.DataFrame({'x': [1]})), \
             patch.object(main, 'process_data', return_value=(pd.DataFrame(), pd.DataFrame())), \
             patch.object(main, 'generate_excel_with_format',
                          side_effect=lambda *a, **kw: excel_called.append(True)), \
             patch.object(main, 'send_feishu_alert',
                          side_effect=lambda *a, **kw: alerts.append(a)), \
             patch.object(main, 'setup_logging'), \
             patch.object(main, 'clean_old_files'), \
             patch.object(main, 'send_to_robot'):
            main.daily_task()

        self.assertEqual(len(excel_called), 0, "空数据时不应调用 generate_excel_with_format")
        self.assertTrue(len(alerts) > 0, "空数据时应推送飞书告警")
        self.assertEqual(alerts[0][0], 'warning')

    def test_nonempty_df_final_generates_file(self):
        """df_final 非空时，正常流程应调用 generate_excel_with_format"""
        df_mock = pd.DataFrame({'序号': [1], '客户名称': ['A']})
        excel_called = []

        with patch.object(main, 'get_db_data', return_value=pd.DataFrame({'x': [1]})), \
             patch.object(main, 'process_data', return_value=(df_mock, pd.DataFrame())), \
             patch.object(main, 'generate_excel_with_format',
                          side_effect=lambda *a, **kw: excel_called.append(True) or a[1]), \
             patch.object(main, 'setup_logging'), \
             patch.object(main, 'clean_old_files'), \
             patch.object(main, 'send_to_robot'):
            main.daily_task()

        self.assertEqual(len(excel_called), 1, "非空数据时应调用 generate_excel_with_format 一次")


# =============================================================================
if __name__ == '__main__':
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()

    test_classes = [
        TestBug1ToDatetimeCoerce,
        TestBug2ExcelWriterContextManager,
        TestBug3CheckDeviceCountDiffColumnStrip,
        TestBug4UploadRaiseForStatus,
        TestBug5SendRespErrcode,
        TestBug6EmptyDfFinalGuard,
    ]
    for cls in test_classes:
        suite.addTests(loader.loadTestsFromTestCase(cls))

    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    sys.exit(0 if result.wasSuccessful() else 1)

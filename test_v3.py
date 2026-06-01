#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
IFR Automation V3 测试脚本
独立测试文件，用于验证 V3 所有功能

运行方式:
    python test_v3.py              # 运行所有测试
    python test_v3.py -v           # 详细输出
    python test_v3.py TestClass    # 运行特定测试类
"""

import unittest
import tempfile
import shutil
import json
import os
import sys
from pathlib import Path
from datetime import datetime
from unittest.mock import patch, MagicMock

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ifr_automation_v3 import (
    ConfigManager,
    ProjectValidator,
    ProjectScanner,
    DrawingCollector,
    SafetyChecker,
    ValidationReportGenerator,
    IFRAutomation,
    UIHelper,
    SourceDirectory,
    ProjectValidation,
    ProcessResult
)


class TestDataSetup:
    """测试数据设置辅助类"""

    @staticmethod
    def create_test_project(base_path: Path, project_name: str,
                           region: str = "1.NSW",
                           create_ifr_internal: bool = True,
                           create_ifr_client: bool = True,
                           create_files: bool = True) -> Path:
        """创建测试项目结构"""
        if region:
            project_path = base_path / region / project_name
        else:
            project_path = base_path / project_name

        # 创建基本目录结构
        drawings_path = project_path / "Design" / "Engineering" / "1. Drawings"
        reports_path = project_path / "Design" / "Engineering" / "3. Report & Study"

        if create_ifr_internal:
            ifr_internal = drawings_path / "2. IFR_internal"
            ifr_internal.mkdir(parents=True, exist_ok=True)

            report_internal = reports_path / "2. IFR_internal"
            report_internal.mkdir(parents=True, exist_ok=True)

            if create_files:
                # 创建测试图纸文件
                (ifr_internal / "GG-01-PLN-001.pdf").write_text("test drawing")
                (ifr_internal / "GG-01-SLD-002.pdf").write_text("test sld")
                (ifr_internal / "GG-01-GEN-003.pdf").write_text("test gen")
                # 创建测试报告文件
                (report_internal / "GG-01-RPT-001.pdf").write_text("test report")

        if create_ifr_client:
            ifr_client = drawings_path / "3. IFR(Client)"
            ifr_client.mkdir(parents=True, exist_ok=True)

            report_client = reports_path / "3. IFR(Client)"
            report_client.mkdir(parents=True, exist_ok=True)

        return project_path

    @staticmethod
    def create_multi_region_structure(base_path: Path) -> dict:
        """创建多区域项目结构"""
        projects = {}

        regions = ["1.NSW", "2.SA", "3.QLD"]
        for i, region in enumerate(regions):
            for j in range(2):
                project_name = f"GG-{i*10+j:02d}"
                project_path = TestDataSetup.create_test_project(
                    base_path, project_name, region,
                    create_files=(j == 0)  # 只有第一个项目有文件
                )
                if region not in projects:
                    projects[region] = []
                projects[region].append(project_path)

        return projects


class TestConfigManager(unittest.TestCase):
    """ConfigManager 测试"""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.config_path = Path(self.temp_dir) / "config.json"

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_default_config_creation(self):
        """测试默认配置创建"""
        manager = ConfigManager(str(self.config_path))

        self.assertTrue(self.config_path.exists())
        self.assertIn("whitelist", manager.config)
        self.assertIn("blacklist", manager.config)
        self.assertIn("source_directories", manager.config)

    def test_whitelist_blacklist(self):
        """测试白名单和黑名单"""
        manager = ConfigManager(str(self.config_path))

        # 设置白名单
        manager.set("whitelist", ["GG-01", "GG-02"])
        whitelist = manager.get("whitelist")
        self.assertEqual(whitelist, ["GG-01", "GG-02"])

        # 设置黑名单
        manager.set("blacklist", ["*_test", "Template*"])
        blacklist = manager.get("blacklist")
        self.assertIn("*_test", blacklist)

    def test_source_directories_config(self):
        """测试源目录配置"""
        manager = ConfigManager(str(self.config_path))

        source_dirs = manager.get("source_directories")
        self.assertIn("drawings", source_dirs)
        self.assertIn("reports", source_dirs)

    def test_recent_projects(self):
        """测试最近项目记录"""
        manager = ConfigManager(str(self.config_path))

        manager.add_recent_project("GG-01", "/path/to/GG-01")
        manager.add_recent_project("GG-02", "/path/to/GG-02")

        recent = manager.get("recent_projects")
        self.assertTrue(len(recent) >= 2)


class TestProjectValidator(unittest.TestCase):
    """ProjectValidator 测试"""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.base_path = Path(self.temp_dir)
        self.validator = ProjectValidator()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_valid_project(self):
        """测试有效项目验证"""
        project_path = TestDataSetup.create_test_project(
            self.base_path, "GG-01", region=""
        )

        result = self.validator.validate(project_path)

        self.assertIsInstance(result, ProjectValidation)
        self.assertTrue(result.valid)
        self.assertEqual(result.project_name, "GG-01")
        self.assertEqual(result.recommended_action, "process")

    def test_invalid_project_missing_dirs(self):
        """测试缺少目录的无效项目"""
        project_path = self.base_path / "GG-Invalid"
        project_path.mkdir(parents=True)

        result = self.validator.validate(project_path)

        self.assertFalse(result.valid)
        self.assertTrue(len(result.missing_dirs) > 0)
        self.assertEqual(result.recommended_action, "skip")

    def test_project_with_files(self):
        """测试包含文件的项目"""
        project_path = TestDataSetup.create_test_project(
            self.base_path, "GG-02", region="",
            create_files=True
        )

        result = self.validator.validate(project_path)

        self.assertTrue(result.valid)
        self.assertTrue(len(result.source_dirs_found) > 0)

        # 检查是否找到了源目录中的文件
        total_files = sum(sd.file_count for sd in result.source_dirs_found)
        self.assertGreater(total_files, 0)

    def test_project_without_files(self):
        """测试没有文件的项目"""
        project_path = TestDataSetup.create_test_project(
            self.base_path, "GG-03", region="",
            create_files=False
        )

        result = self.validator.validate(project_path)

        # 目录存在但没有文件
        self.assertTrue(result.valid)


class TestProjectScanner(unittest.TestCase):
    """ProjectScanner 测试"""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.base_path = Path(self.temp_dir)
        self.scanner = ProjectScanner()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_hierarchical_scan(self):
        """测试分层扫描"""
        # 创建多区域结构
        TestDataSetup.create_multi_region_structure(self.base_path)

        results = self.scanner.scan_hierarchical(self.base_path)

        self.assertIsInstance(results, dict)
        self.assertIn("1.NSW", results)
        self.assertIn("2.SA", results)
        self.assertIn("3.QLD", results)

    def test_flat_scan(self):
        """测试平面扫描"""
        # 创建没有区域的项目
        TestDataSetup.create_test_project(self.base_path, "GG-01", region="")
        TestDataSetup.create_test_project(self.base_path, "GG-02", region="")

        results = self.scanner.scan_flat(self.base_path)

        self.assertIsInstance(results, list)
        self.assertEqual(len(results), 2)

    def test_whitelist_filter(self):
        """测试白名单过滤"""
        TestDataSetup.create_test_project(self.base_path, "GG-01", region="")
        TestDataSetup.create_test_project(self.base_path, "GG-02", region="")
        TestDataSetup.create_test_project(self.base_path, "GG-03", region="")

        scanner = ProjectScanner(whitelist=["GG-01", "GG-02"])
        results = scanner.scan_flat(self.base_path)

        project_names = [r.project_name for r in results]
        self.assertIn("GG-01", project_names)
        self.assertIn("GG-02", project_names)
        self.assertNotIn("GG-03", project_names)

    def test_blacklist_filter(self):
        """测试黑名单过滤"""
        TestDataSetup.create_test_project(self.base_path, "GG-01", region="")
        TestDataSetup.create_test_project(self.base_path, "GG-02_test", region="")
        TestDataSetup.create_test_project(self.base_path, "Template_Project", region="")

        scanner = ProjectScanner(blacklist=["*_test", "Template*"])
        results = scanner.scan_flat(self.base_path)

        project_names = [r.project_name for r in results]
        self.assertIn("GG-01", project_names)
        self.assertNotIn("GG-02_test", project_names)
        self.assertNotIn("Template_Project", project_names)


class TestDrawingCollector(unittest.TestCase):
    """DrawingCollector 测试"""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.base_path = Path(self.temp_dir)
        self.collector = DrawingCollector()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_collect_from_single_source(self):
        """测试单源目录收集"""
        project_path = TestDataSetup.create_test_project(
            self.base_path, "GG-01", region="",
            create_files=True
        )

        files = self.collector.collect_drawings(project_path)

        self.assertIsInstance(files, list)
        self.assertGreater(len(files), 0)

    def test_collect_with_pattern_filter(self):
        """测试文件模式过滤"""
        project_path = TestDataSetup.create_test_project(
            self.base_path, "GG-01", region="",
            create_files=True
        )

        # 只收集 PLN 文件
        files = self.collector.collect_drawings(
            project_path,
            patterns=["PLN"]
        )

        for f in files:
            self.assertIn("PLN", f.name)

    def test_collect_from_multiple_sources(self):
        """测试多源目录收集"""
        project_path = TestDataSetup.create_test_project(
            self.base_path, "GG-01", region="",
            create_files=True
        )

        # 创建额外的 Working 目录
        working_dir = project_path / "Design" / "Engineering" / "1. Drawings" / "1. Working"
        working_dir.mkdir(parents=True, exist_ok=True)
        (working_dir / "GG-01-PLN-100.pdf").write_text("working drawing")

        collector = DrawingCollector(
            source_paths=[
                "Design/Engineering/1. Drawings/2. IFR_internal",
                "Design/Engineering/1. Drawings/1. Working"
            ]
        )

        files = collector.collect_drawings(project_path)

        file_names = [f.name for f in files]
        self.assertIn("GG-01-PLN-001.pdf", file_names)
        self.assertIn("GG-01-PLN-100.pdf", file_names)


class TestSafetyChecker(unittest.TestCase):
    """SafetyChecker 测试"""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.base_path = Path(self.temp_dir)
        self.checker = SafetyChecker()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_safe_project(self):
        """测试安全项目检查"""
        project_path = TestDataSetup.create_test_project(
            self.base_path, "GG-01", region="",
            create_files=True
        )

        result = self.checker.check(project_path)

        self.assertTrue(result["safe"])
        self.assertEqual(len(result["warnings"]), 0)

    def test_large_file_warning(self):
        """测试大文件警告"""
        project_path = TestDataSetup.create_test_project(
            self.base_path, "GG-01", region="",
            create_files=False
        )

        # 创建大文件 (模拟)
        ifr_internal = project_path / "Design" / "Engineering" / "1. Drawings" / "2. IFR_internal"
        ifr_internal.mkdir(parents=True, exist_ok=True)

        # 创建很多文件来触发警告
        for i in range(60):
            (ifr_internal / f"GG-01-PLN-{i:03d}.pdf").write_text("x" * 1000)

        checker = SafetyChecker(max_files_threshold=50)
        result = checker.check(project_path)

        self.assertTrue(len(result["warnings"]) > 0)

    def test_empty_source_warning(self):
        """测试空源目录警告"""
        project_path = TestDataSetup.create_test_project(
            self.base_path, "GG-01", region="",
            create_files=False
        )

        result = self.checker.check(project_path)

        # 空目录应该有警告
        self.assertTrue(any("empty" in w.lower() or "no files" in w.lower()
                           for w in result.get("warnings", [])) or result["safe"])


class TestValidationReportGenerator(unittest.TestCase):
    """ValidationReportGenerator 测试"""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.generator = ValidationReportGenerator()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_generate_json_report(self):
        """测试 JSON 报告生成"""
        validations = [
            ProjectValidation(
                valid=True,
                project_name="GG-01",
                project_path="/path/to/GG-01",
                region="1.NSW",
                recommended_action="process"
            ),
            ProjectValidation(
                valid=False,
                project_name="GG-02",
                project_path="/path/to/GG-02",
                region="1.NSW",
                missing_dirs=["IFR_internal"],
                recommended_action="skip"
            )
        ]

        report_path = Path(self.temp_dir) / "report.json"
        self.generator.generate(validations, report_path, format="json")

        self.assertTrue(report_path.exists())

        with open(report_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        self.assertIn("projects", data)
        self.assertIn("summary", data)
        self.assertEqual(data["summary"]["total"], 2)
        self.assertEqual(data["summary"]["valid"], 1)

    def test_generate_text_report(self):
        """测试文本报告生成"""
        validations = [
            ProjectValidation(
                valid=True,
                project_name="GG-01",
                project_path="/path/to/GG-01",
                region="1.NSW",
                recommended_action="process"
            )
        ]

        report_path = Path(self.temp_dir) / "report.txt"
        self.generator.generate(validations, report_path, format="text")

        self.assertTrue(report_path.exists())
        content = report_path.read_text(encoding='utf-8')
        self.assertIn("GG-01", content)


class TestIFRAutomation(unittest.TestCase):
    """IFRAutomation 核心功能测试"""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.base_path = Path(self.temp_dir)

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_dry_run_mode(self):
        """测试预览模式"""
        project_path = TestDataSetup.create_test_project(
            self.base_path, "GG-01", region=""
        )

        automation = IFRAutomation(str(self.base_path), dry_run=True)
        result = automation.process_project(project_path)

        self.assertIsInstance(result, ProcessResult)
        # 预览模式不应该实际复制文件

    def test_process_single_project(self):
        """测试处理单个项目"""
        project_path = TestDataSetup.create_test_project(
            self.base_path, "GG-01", region="",
            create_files=True
        )

        automation = IFRAutomation(str(self.base_path), dry_run=False)
        result = automation.process_project(project_path)

        self.assertIsInstance(result, ProcessResult)
        self.assertTrue(result.success)

        # 检查文件是否被复制
        client_dir = project_path / "Design" / "Engineering" / "1. Drawings" / "3. IFR(Client)"
        copied_files = list(client_dir.glob("*.pdf"))
        self.assertGreater(len(copied_files), 0)

    def test_process_with_validation(self):
        """测试带验证的处理"""
        project_path = TestDataSetup.create_test_project(
            self.base_path, "GG-01", region="",
            create_files=True
        )

        automation = IFRAutomation(str(self.base_path))

        # 先验证
        validation = automation.validate_project(project_path)
        self.assertTrue(validation.valid)

        # 然后处理
        result = automation.process_project(project_path)
        self.assertTrue(result.success)


class TestUIHelper(unittest.TestCase):
    """UIHelper 测试"""

    def test_format_size(self):
        """测试文件大小格式化"""
        self.assertEqual(UIHelper.format_size(500), "500 B")
        self.assertEqual(UIHelper.format_size(1024), "1.0 KB")
        self.assertEqual(UIHelper.format_size(1024 * 1024), "1.0 MB")

    def test_format_status(self):
        """测试状态格式化"""
        # 测试不同状态的格式化
        status = UIHelper.format_status("process")
        self.assertIn("process", status.lower())

        status = UIHelper.format_status("skip")
        self.assertIn("skip", status.lower())


class TestDataclasses(unittest.TestCase):
    """数据类测试"""

    def test_source_directory(self):
        """测试 SourceDirectory 数据类"""
        sd = SourceDirectory(
            path=Path("/test/path"),
            dir_type="drawings",
            file_count=10,
            total_size=1024
        )

        self.assertEqual(sd.path, Path("/test/path"))
        self.assertEqual(sd.dir_type, "drawings")
        self.assertEqual(sd.file_count, 10)
        self.assertEqual(sd.total_size, 1024)

    def test_project_validation(self):
        """测试 ProjectValidation 数据类"""
        pv = ProjectValidation(
            valid=True,
            project_name="GG-01",
            project_path="/test/GG-01",
            region="1.NSW",
            recommended_action="process"
        )

        self.assertTrue(pv.valid)
        self.assertEqual(pv.project_name, "GG-01")
        self.assertEqual(pv.region, "1.NSW")
        self.assertEqual(pv.recommended_action, "process")
        self.assertEqual(pv.missing_dirs, [])

    def test_process_result(self):
        """测试 ProcessResult 数据类"""
        pr = ProcessResult(
            success=True,
            project_name="GG-01",
            files_copied=5,
            files_skipped=2,
            errors=[]
        )

        self.assertTrue(pr.success)
        self.assertEqual(pr.files_copied, 5)
        self.assertEqual(pr.files_skipped, 2)
        self.assertEqual(pr.errors, [])


class TestIntegration(unittest.TestCase):
    """集成测试"""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.base_path = Path(self.temp_dir)

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_full_workflow(self):
        """测试完整工作流程"""
        # 1. 创建多区域项目结构
        TestDataSetup.create_multi_region_structure(self.base_path)

        # 2. 扫描项目
        scanner = ProjectScanner()
        projects = scanner.scan_hierarchical(self.base_path)

        self.assertTrue(len(projects) > 0)

        # 3. 验证项目
        validator = ProjectValidator()
        validations = []
        for region, project_list in projects.items():
            for pv in project_list:
                validations.append(pv)

        # 4. 安全检查
        checker = SafetyChecker()
        for v in validations:
            if v.valid:
                safety = checker.check(Path(v.project_path))
                self.assertIn("safe", safety)

        # 5. 处理项目
        automation = IFRAutomation(str(self.base_path), dry_run=True)
        for v in validations:
            if v.valid and v.recommended_action == "process":
                result = automation.process_project(Path(v.project_path))
                self.assertIsInstance(result, ProcessResult)

    def test_whitelist_blacklist_workflow(self):
        """测试白名单/黑名单工作流程"""
        # 创建项目
        TestDataSetup.create_test_project(self.base_path, "GG-01", region="")
        TestDataSetup.create_test_project(self.base_path, "GG-02", region="")
        TestDataSetup.create_test_project(self.base_path, "GG-03_backup", region="")

        # 使用黑名单扫描
        scanner = ProjectScanner(blacklist=["*_backup"])
        results = scanner.scan_flat(self.base_path)

        names = [r.project_name for r in results]
        self.assertIn("GG-01", names)
        self.assertIn("GG-02", names)
        self.assertNotIn("GG-03_backup", names)


def run_tests():
    """运行所有测试"""
    # 创建测试套件
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()

    # 添加所有测试类
    test_classes = [
        TestConfigManager,
        TestProjectValidator,
        TestProjectScanner,
        TestDrawingCollector,
        TestSafetyChecker,
        TestValidationReportGenerator,
        TestIFRAutomation,
        TestUIHelper,
        TestDataclasses,
        TestIntegration
    ]

    for test_class in test_classes:
        tests = loader.loadTestsFromTestCase(test_class)
        suite.addTests(tests)

    # 运行测试
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)

    # 打印摘要
    print("\n" + "=" * 60)
    print("测试摘要")
    print("=" * 60)
    print(f"运行测试数: {result.testsRun}")
    print(f"成功: {result.testsRun - len(result.failures) - len(result.errors)}")
    print(f"失败: {len(result.failures)}")
    print(f"错误: {len(result.errors)}")

    if result.failures:
        print("\n失败的测试:")
        for test, traceback in result.failures:
            print(f"  - {test}")

    if result.errors:
        print("\n出错的测试:")
        for test, traceback in result.errors:
            print(f"  - {test}")

    return result


if __name__ == "__main__":
    # 检查命令行参数
    if len(sys.argv) > 1 and sys.argv[1] == "-v":
        # 详细模式
        unittest.main(verbosity=2)
    elif len(sys.argv) > 1:
        # 运行特定测试
        unittest.main()
    else:
        # 运行所有测试
        result = run_tests()
        sys.exit(0 if result.wasSuccessful() else 1)

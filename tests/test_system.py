"""
焊接领域专有名词提取系统测试
"""

import unittest
import os
import sys

# Add src directory to Python path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from core.engine import ExtractionEngine
from processors.result_processor import ResultProcessor
from parsers.document_parser import DocumentParser
from extractors import ExtractionResult


class TestExtractionEngine(unittest.TestCase):
    """测试提取引擎"""
    
    def setUp(self):
        """设置测试环境"""
        self.engine = ExtractionEngine()
        self.processor = ResultProcessor()
    
    def test_engine_initialization(self):
        """测试引擎初始化"""
        self.assertIsNotNone(self.engine)
        self.assertEqual(set(self.engine.extractor_paths), {'api', 'ollama'})
        self.assertNotIn('jieba', self.engine.extractor_paths)
    
    def test_fallback_mechanism(self):
        """测试降级机制"""
        text = "焊接工艺包括TIG焊、MIG焊、CO2焊等多种方法。"
        
        result_data = self.engine.extract_with_fallback(text)
        
        self.assertIn('success', result_data)
        self.assertIn('terms', result_data)
        self.assertIn('source', result_data)


class TestResultProcessor(unittest.TestCase):
    """测试结果处理器"""
    
    def setUp(self):
        """设置测试环境"""
        self.processor = ResultProcessor()
    
    def test_normalize_term(self):
        """测试术语标准化"""
        self.assertEqual(self.processor.normalize_term('氩弧焊'), 'TIG焊')
        self.assertEqual(self.processor.normalize_term('  TIG焊  '), 'TIG焊')
    
    def test_deduplicate(self):
        """测试去重"""
        results = [
            ExtractionResult(term='TIG焊', category='法', source='api'),
            ExtractionResult(term='TIG焊', category='法', source='ollama'),
            ExtractionResult(term='MIG焊', category='法', source='api')
        ]
        
        deduplicated = self.processor.deduplicate(results)
        
        self.assertLess(len(deduplicated), len(results))
    
    def test_process(self):
        """测试完整处理流程"""
        results = [
            ExtractionResult(term='TIG焊', source='api'),
            ExtractionResult(term='MIG焊', source='api'),
            ExtractionResult(term='焊接电流', source='ollama')
        ]
        
        processed = self.processor.process(results)
        
        self.assertIsInstance(processed, list)
        self.assertGreater(len(processed), 0)


class TestDocumentParser(unittest.TestCase):
    """测试文档解析器"""
    
    def setUp(self):
        """设置测试环境"""
        self.parser = DocumentParser()
    
    def test_txt_parsing(self):
        """测试TXT解析"""
        test_file = 'test_document.txt'
        
        try:
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write('TIG焊是一种常用的焊接方法。')
            
            content = self.parser.parse(test_file)
            
            self.assertIn('TIG焊', content)
            
        finally:
            if os.path.exists(test_file):
                os.remove(test_file)
    
    def test_supported_formats(self):
        """测试支持的格式"""
        formats = self.parser.get_supported_formats()
        
        self.assertIn('.txt', formats)
        self.assertIn('.pdf', formats)
        self.assertIn('.docx', formats)


class TestIntegration(unittest.TestCase):
    """集成测试"""
    
    def setUp(self):
        """设置测试环境"""
        self.engine = ExtractionEngine()
        self.processor = ResultProcessor()
        self.parser = DocumentParser()
    
    def test_full_extraction_pipeline(self):
        """测试完整提取流程"""
        text = """
        焊接工艺是现代制造业的重要组成部分。TIG焊（钨极氩弧焊）是一种高质量的焊接方法，
        广泛应用于不锈钢、铝合金等材料的焊接。MIG焊和CO2焊也是常用的气体保护焊方法。
        在焊接过程中，需要严格控制焊接电流、焊接电压、焊接速度等工艺参数。
        焊接缺陷如气孔、裂纹、夹渣等会影响焊接质量，因此需要进行无损检测。
        """
        
        if not self.engine.get_available_extractors():
            self.skipTest('API/Ollama 提取器不可用，跳过外部服务集成测试')
        result_data = self.engine.extract_with_fallback(text)
        
        self.assertTrue(result_data['success'])
        self.assertGreater(len(result_data['terms']), 0)
        
        from extractors import ExtractionResult
        raw_results = [ExtractionResult(**term) for term in result_data['terms']]
        processed_results = self.processor.process(raw_results)
        
        self.assertGreater(len(processed_results), 0)
        
        statistics = self.processor.get_statistics(processed_results)
        self.assertIn('total_terms', statistics)
        self.assertIn('categories', statistics)


def run_tests():
    """运行所有测试"""
    unittest.main(argv=[''], verbosity=2, exit=False)


if __name__ == '__main__':
    run_tests()

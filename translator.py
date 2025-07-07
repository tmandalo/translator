"""
Основной модуль переводчика с OpenRouter API
"""

import time
import json
import logging
from typing import List, Dict, Any, Optional
from dataclasses import dataclass

import requests
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

from config import config
from text_chunker import TextChunk


@dataclass
class TranslationResult:
    """Результат перевода"""
    original_text: str
    translated_text: str
    success: bool
    error: Optional[str] = None
    tokens_used: Optional[int] = None
    processing_time: Optional[float] = None


class OpenRouterTranslator:
    """Класс для работы с OpenRouter API"""
    
    def __init__(self):
        self.api_key = config.openrouter_api_key
        self.model = config.openrouter_model
        self.base_url = "https://openrouter.ai/api/v1/chat/completions"
        self.session = requests.Session()
        
        # Настраиваем заголовки
        self.session.headers.update({
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
            "HTTP-Referer": "https://github.com/your-repo",  # Обязательно для OpenRouter
            "X-Title": "Literary Document Translator"
        })
        
        # Настраиваем логирование
        self.logger = logging.getLogger(__name__)
        
    def get_translation_prompt(self) -> str:
        """ФИНАЛЬНЫЙ ПРОМПТ: Баланс литературности, точности и аутентичности."""
        return """
Вы — элитный литературный переводчик с английского на русский. Ваша задача — перевести текст так, чтобы он читался как произведение, изначально написанное на русском языке, но при этом с абсолютной точностью передавал все детали, стиль и голос оригинального автора.

ОСНОВНЫЕ ПРИНЦИПЫ:

1.  **ВЕРНОСТЬ АВТОРУ:**
    -   **Стиль и Тон:** Полностью сохраняйте уникальный авторский стиль: его ритм, интонацию, юмор, сарказм или серьезность.
    -   **Точность деталей:** Не упускайте ни одной детали, метафоры или нюанса. Ваш перевод должен быть исчерпывающим. Обобщения и сокращения строго запрещены.
    -   **АУТЕНТИЧНОСТЬ ЛЕКСИКИ:** Сохраняйте всю лексику автора, включая разговорные выражения, сленг и нецензурную брань. Цель — полная аутентичность, а не стерильность текста. Не заменяйте грубые слова эвфемизмами.

2.  **НАТИВНОСТЬ ДЛЯ ЧИТАТЕЛЯ:**
    -   **Живой язык:** Используйте богатый и естественный русский язык. Перевод не должен звучать как подстрочник.
    -   **Адаптация идиом:** Адаптируйте идиомы и фразеологизмы так, чтобы они были понятны русскоязычному читателю, сохраняя при этом первоначальный смысл и эффект.

ВАШЕ КРЕДО: "Точность оригинала, воплощенная в красоте родного языка". Вы не упрощаете и не додумываете. Вы пересоздаете.

Переведите следующий текст, строго придерживаясь этих принципов:
"""
    
    def _calculate_optimal_max_tokens(self, text: str) -> int:
        """
        Вычисляет оптимальное количество токенов для перевода на основе размера входного текста
        
        Args:
            text: Входной текст
            
        Returns:
            Оптимальное количество max_tokens
        """
        # Приблизительная оценка: 1 токен ≈ 4 символа для английского
        # Для русского обычно нужно больше токенов (коэффициент 1.2-1.5)
        input_tokens_estimate = len(text) // 4
        output_tokens_estimate = int(input_tokens_estimate * 1.3)  # Коэффициент для русского
        
        # Используем настройку из конфигурации как максимум
        max_allowed = getattr(config, 'max_tokens', 15000)
        
        # Минимум 2000 токенов для коротких текстов
        min_tokens = 2000
        
        # Оптимальное значение с ограничениями
        optimal_tokens = max(min_tokens, min(output_tokens_estimate, max_allowed))
        
        self.logger.debug(f"Текст: {len(text)} символов, расчетные токены вывода: {output_tokens_estimate}, используем: {optimal_tokens}")
        
        return optimal_tokens
    
    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(multiplier=1, min=1, max=10),
        retry=retry_if_exception_type((requests.exceptions.RequestException, requests.exceptions.Timeout))
    )
    def translate_text(self, text: str) -> TranslationResult:
        """
        Переводит текст с английского на русский
        
        Args:
            text: Текст для перевода
            
        Returns:
            Результат перевода
        """
        start_time = time.time()
        
        try:
            # Подготавливаем данные для запроса
            payload = {
                "model": self.model,
                "messages": [
                    {
                        "role": "system",
                        "content": self.get_translation_prompt()
                    },
                    {
                        "role": "user",
                        "content": text
                    }
                ],
                "max_tokens": self._calculate_optimal_max_tokens(text),
                "temperature": 0.3,
                "top_p": 0.9,
                "frequency_penalty": 0.0,
                "presence_penalty": 0.0
            }
            
            # Отправляем запрос
            response = self.session.post(
                self.base_url,
                json=payload,
                timeout=config.request_timeout
            )
            
            # Проверяем статус ответа
            response.raise_for_status()
            
            # Парсим ответ
            response_data = response.json()
            
            if 'choices' not in response_data or not response_data['choices']:
                raise ValueError("Некорректный ответ от API - нет choices")
                
            translated_text = response_data['choices'][0]['message']['content'].strip()
            
            # Извлекаем количество токенов если доступно
            tokens_used = None
            if 'usage' in response_data:
                tokens_used = response_data['usage'].get('total_tokens')
            
            processing_time = time.time() - start_time
            
            self.logger.info(f"Перевод выполнен за {processing_time:.2f}с, токенов: {tokens_used}")
            
            return TranslationResult(
                original_text=text,
                translated_text=translated_text,
                success=True,
                tokens_used=tokens_used,
                processing_time=processing_time
            )
            
        except requests.exceptions.RequestException as e:
            error_msg = f"Ошибка сети: {str(e)}"
            self.logger.error(error_msg)
            return TranslationResult(
                original_text=text,
                translated_text="",
                success=False,
                error=error_msg,
                processing_time=time.time() - start_time
            )
            
        except json.JSONDecodeError as e:
            error_msg = f"Ошибка парсинга JSON: {str(e)}"
            self.logger.error(error_msg)
            return TranslationResult(
                original_text=text,
                translated_text="",
                success=False,
                error=error_msg,
                processing_time=time.time() - start_time
            )
            
        except Exception as e:
            error_msg = f"Неожиданная ошибка: {str(e)}"
            self.logger.error(error_msg)
            return TranslationResult(
                original_text=text,
                translated_text="",
                success=False,
                error=error_msg,
                processing_time=time.time() - start_time
            )
    
    def translate_chunks(self, chunks: List[TextChunk], progress_callback=None) -> List[TranslationResult]:
        """
        Переводит список блоков текста (оптимизировано для больших блоков)
        
        Args:
            chunks: Список блоков для перевода
            progress_callback: Функция для отслеживания прогресса
            
        Returns:
            Список результатов перевода
        """
        results = []
        total_chars = sum(len(chunk.text) for chunk in chunks)
        processed_chars = 0
        
        self.logger.info(f"🚀 Начинаем последовательный перевод {len(chunks)} блоков ({total_chars:,} символов)")
        
        for i, chunk in enumerate(chunks):
            chunk_size = len(chunk.text)
            self.logger.info(f"Переводим блок {i+1}/{len(chunks)} ({chunk_size:,} символов)")
            
            # Переводим блок
            result = self.translate_text(chunk.text)
            results.append(result)
            
            # Обновляем прогресс
            processed_chars += chunk_size
            if progress_callback:
                progress_callback(i + 1, len(chunks), result.success)
            
            # Логируем промежуточную статистику
            if result.success:
                tokens_used = result.tokens_used or 0
                efficiency = tokens_used / chunk_size if chunk_size > 0 else 0
                self.logger.info(f"✅ Блок {i+1} переведен: {tokens_used} токенов, эффективность: {efficiency:.3f} токен/символ")
            else:
                self.logger.warning(f"❌ Ошибка блока {i+1}: {result.error}")
        
        # Финальная статистика
        successful = len([r for r in results if r.success])
        total_tokens = sum(r.tokens_used for r in results if r.tokens_used)
        self.logger.info(f"🎉 Перевод завершен: {successful}/{len(chunks)} блоков, {total_tokens:,} токенов")
        
        return results
    
    def get_translation_statistics(self, results: List[TranslationResult]) -> Dict[str, Any]:
        """Возвращает статистику перевода"""
        if not results:
            return {}
        
        successful = [r for r in results if r.success]
        failed = [r for r in results if not r.success]
        
        total_tokens = sum(r.tokens_used for r in successful if r.tokens_used)
        total_time = sum(r.processing_time for r in results if r.processing_time)
        
        return {
            'total_chunks': len(results),
            'successful_chunks': len(successful),
            'failed_chunks': len(failed),
            'success_rate': len(successful) / len(results) if results else 0,
            'total_tokens_used': total_tokens,
            'total_processing_time': total_time,
            'average_chunk_time': total_time / len(results) if results else 0,
            'errors': [r.error for r in failed if r.error]
        }
    
    def test_connection(self) -> bool:
        """Тестирует соединение с API"""
        try:
            test_result = self.translate_text("Hello, world!")
            return test_result.success
        except Exception as e:
            self.logger.error(f"Ошибка тестирования соединения: {e}")
            return False


class DocumentTranslator:
    """Высокоуровневый класс для перевода документов"""
    
    def __init__(self):
        self.api_translator = OpenRouterTranslator()
        self.logger = logging.getLogger(__name__)
    
    def translate_document_chunks(self, chunks: List[TextChunk], progress_callback=None) -> List[TranslationResult]:
        """
        Переводит блоки документа
        
        Args:
            chunks: Список блоков текста
            progress_callback: Функция для отслеживания прогресса
            
        Returns:
            Список результатов перевода
        """
        self.logger.info(f"Начинаем перевод {len(chunks)} блоков")
        
        # Проверяем соединение с API
        if not self.api_translator.test_connection():
            self.logger.error("Не удалось подключиться к API")
            return []
        
        # Переводим блоки
        results = self.api_translator.translate_chunks(chunks, progress_callback)
        
        # Выводим статистику
        stats = self.api_translator.get_translation_statistics(results)
        self.logger.info(f"Перевод завершен. Успешно: {stats['successful_chunks']}/{stats['total_chunks']}")
        
        return results 
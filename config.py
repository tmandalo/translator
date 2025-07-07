"""
Конфигурация для литературного переводчика документов
"""

import os
from typing import Optional
from dotenv import load_dotenv
from pydantic_settings import BaseSettings
from pydantic import Field

# Загружаем переменные окружения
load_dotenv()

class Config(BaseSettings):
    """Конфигурация приложения"""
    
    # OpenRouter API
    openrouter_api_key: str = Field(..., env='OPENROUTER_API_KEY')
    openrouter_model: str = Field('openrouter/gpt-4.1-nano', env='OPENROUTER_MODEL')
    
    # Настройки перевода (оптимизировано под 128K context)
    chunk_size: int = Field(45000, env='CHUNK_SIZE')  # ~11K токенов
    max_retries: int = Field(3, env='MAX_RETRIES')
    retry_delay: float = Field(1.0, env='RETRY_DELAY')
    request_timeout: float = Field(180.0, env='REQUEST_TIMEOUT')  # Увеличен таймаут для больших блоков
    max_tokens: int = Field(15000, env='MAX_TOKENS')  # Близко к лимиту 16K
    
    # 🚀 Настройки оптимизации (максимальное использование 128K context)
    enable_optimization: bool = Field(True, env='ENABLE_OPTIMIZATION')
    batch_size: int = Field(3, env='BATCH_SIZE')  # Уменьшено из-за больших блоков
    max_concurrent_requests: int = Field(2, env='MAX_CONCURRENT_REQUESTS')  # Уменьшено для стабильности
    optimal_chunk_size: int = Field(80000, env='OPTIMAL_CHUNK_SIZE')  # ~20K токенов для batch
    use_async: bool = Field(True, env='USE_ASYNC')  # Использовать асинхронные запросы
    
    # Настройки вывода
    save_xml: bool = Field(False, env='SAVE_XML')
    log_level: str = Field('INFO', env='LOG_LEVEL')
    
    class Config:
        env_file = '.env'
        case_sensitive = False

# Глобальный экземпляр конфигурации
config = Config() 
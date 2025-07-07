"""
Модуль для разбивки текста на смысловые блоки
"""

import re
from typing import List, Tuple
from dataclasses import dataclass

@dataclass
class TextChunk:
    """Класс для хранения блока текста с метаданными"""
    text: str
    start_index: int
    end_index: int
    paragraph_index: int
    is_complete_paragraph: bool


class TextChunker:
    """Класс для разбивки текста на оптимальные блоки для перевода"""
    
    def __init__(self, max_chunk_size: int = 45000):
        self.max_chunk_size = max_chunk_size
        
    def chunk_text(self, text: str) -> List[TextChunk]:
        """
        Разбивает текст на смысловые блоки
        
        Args:
            text: Исходный текст для разбивки
            
        Returns:
            Список блоков текста
        """
        if not text.strip():
            return []
            
        # Разбиваем на параграфы
        paragraphs = self._split_into_paragraphs(text)
        chunks = []
        current_chunk = ""
        current_start = 0
        paragraph_indices = []
        
        for i, paragraph in enumerate(paragraphs):
            # Если параграф сам по себе больше лимита, разбиваем его
            if len(paragraph) > self.max_chunk_size:
                # Сохраняем текущий чанк если он не пустой
                if current_chunk.strip():
                    chunks.append(TextChunk(
                        text=current_chunk.strip(),
                        start_index=current_start,
                        end_index=current_start + len(current_chunk),
                        paragraph_index=paragraph_indices[0] if paragraph_indices else i,
                        is_complete_paragraph=len(paragraph_indices) == 1
                    ))
                
                # Разбиваем большой параграф на предложения
                long_paragraph_chunks = self._split_long_paragraph(paragraph, i)
                chunks.extend(long_paragraph_chunks)
                
                # Обновляем позиции
                current_start += len(current_chunk) + len(paragraph)
                current_chunk = ""
                paragraph_indices = []
                
            # Если добавление параграфа не превышает лимит
            elif len(current_chunk) + len(paragraph) <= self.max_chunk_size:
                current_chunk += paragraph
                paragraph_indices.append(i)
                
            # Если добавление параграфа превышает лимит
            else:
                # Сохраняем текущий чанк
                if current_chunk.strip():
                    chunks.append(TextChunk(
                        text=current_chunk.strip(),
                        start_index=current_start,
                        end_index=current_start + len(current_chunk),
                        paragraph_index=paragraph_indices[0] if paragraph_indices else i,
                        is_complete_paragraph=len(paragraph_indices) == 1
                    ))
                
                # Начинаем новый чанк
                current_start += len(current_chunk)
                current_chunk = paragraph
                paragraph_indices = [i]
        
        # Добавляем последний чанк
        if current_chunk.strip():
            chunks.append(TextChunk(
                text=current_chunk.strip(),
                start_index=current_start,
                end_index=current_start + len(current_chunk),
                paragraph_index=paragraph_indices[0] if paragraph_indices else len(paragraphs) - 1,
                is_complete_paragraph=len(paragraph_indices) == 1
            ))
        
        return chunks
    
    def _split_into_paragraphs(self, text: str) -> List[str]:
        """Разбивает текст на параграфы"""
        # Разбиваем по двойным переносам строк
        paragraphs = re.split(r'\n\s*\n', text)
        
        # Очищаем параграфы от лишних пробелов, но сохраняем структуру
        cleaned_paragraphs = []
        for paragraph in paragraphs:
            cleaned = paragraph.strip()
            if cleaned:
                cleaned_paragraphs.append(cleaned + '\n\n')
        
        return cleaned_paragraphs
    
    def _split_long_paragraph(self, paragraph: str, paragraph_index: int) -> List[TextChunk]:
        """Разбивает слишком длинный параграф на предложения"""
        # Разбиваем на предложения
        sentences = re.split(r'(?<=[.!?])\s+', paragraph)
        
        chunks = []
        current_chunk = ""
        start_pos = 0
        
        for sentence in sentences:
            if len(current_chunk) + len(sentence) <= self.max_chunk_size:
                current_chunk += sentence + " "
            else:
                if current_chunk.strip():
                    chunks.append(TextChunk(
                        text=current_chunk.strip(),
                        start_index=start_pos,
                        end_index=start_pos + len(current_chunk),
                        paragraph_index=paragraph_index,
                        is_complete_paragraph=False
                    ))
                
                start_pos += len(current_chunk)
                current_chunk = sentence + " "
        
        # Добавляем последний кусок
        if current_chunk.strip():
            chunks.append(TextChunk(
                text=current_chunk.strip(),
                start_index=start_pos,
                end_index=start_pos + len(current_chunk),
                paragraph_index=paragraph_index,
                is_complete_paragraph=False
            ))
        
        return chunks
    
    def get_chunk_statistics(self, chunks: List[TextChunk]) -> dict:
        """Возвращает статистику по блокам"""
        if not chunks:
            return {
                'total_chunks': 0,
                'total_characters': 0,
                'average_chunk_size': 0,
                'max_chunk_size': 0,
                'min_chunk_size': 0,
                'complete_paragraphs': 0
            }
        
        sizes = [len(chunk.text) for chunk in chunks]
        
        return {
            'total_chunks': len(chunks),
            'total_characters': sum(sizes),
            'average_chunk_size': sum(sizes) / len(sizes),
            'max_chunk_size': max(sizes),
            'min_chunk_size': min(sizes),
            'complete_paragraphs': sum(1 for chunk in chunks if chunk.is_complete_paragraph)
        } 
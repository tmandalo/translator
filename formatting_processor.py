#!/usr/bin/env python3
"""
Модуль для продвинутой обработки форматирования документов
"""

import re
from typing import List, Dict, Any, Tuple, Optional
from dataclasses import dataclass
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_COLOR_INDEX, WD_UNDERLINE


@dataclass
class FormattingSegment:
    """Сегмент текста с форматированием"""
    text: str
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    font_name: Optional[str] = None
    font_size: Optional[float] = None
    font_color: Optional[str] = None
    start_pos: int = 0
    end_pos: int = 0


class FormattingProcessor:
    """Класс для обработки форматирования при переводе"""
    
    def __init__(self):
        pass
    
    def extract_formatting_segments(self, original_text: str, formatting_data: Dict[str, Any]) -> List[FormattingSegment]:
        """
        Извлекает сегменты форматирования из оригинального текста
        
        Args:
            original_text: Оригинальный текст
            formatting_data: Данные форматирования из extract_paragraph_formatting
            
        Returns:
            Список сегментов с форматированием
        """
        if not formatting_data or 'runs' not in formatting_data:
            # Если нет данных форматирования, возвращаем весь текст как один сегмент
            return [FormattingSegment(
                text=original_text,
                start_pos=0,
                end_pos=len(original_text)
            )]
        
        segments = []
        current_pos = 0
        
        for run_data in formatting_data['runs']:
            run_text = run_data.get('text', '')
            
            if not run_text:
                continue
                
            # Находим позицию этого run в общем тексте
            text_start = original_text.find(run_text, current_pos)
            if text_start == -1:
                # Если не можем найти точное совпадение, используем текущую позицию
                text_start = current_pos
            
            text_end = text_start + len(run_text)
            
            # Создаем сегмент форматирования
            segment = FormattingSegment(
                text=run_text,
                bold=run_data.get('bold'),
                italic=run_data.get('italic'),
                underline=run_data.get('underline'),
                font_name=run_data.get('font_name'),
                font_size=self._convert_font_size(run_data.get('font_size')),
                font_color=self._convert_font_color(run_data.get('font_color')),
                start_pos=text_start,
                end_pos=text_end
            )
            
            segments.append(segment)
            current_pos = text_end
        
        return segments
    
    def _convert_font_size(self, font_size) -> Optional[float]:
        """Конвертирует размер шрифта в точки"""
        if font_size is None:
            return None
        
        # font_size может быть в разных единицах, обычно это Pt объект
        try:
            if hasattr(font_size, 'pt'):
                return float(font_size.pt)
            else:
                return float(font_size) if font_size else None
        except (TypeError, ValueError):
            return None
    
    def _convert_font_color(self, font_color) -> Optional[str]:
        """Конвертирует цвет шрифта в hex формат"""
        if font_color is None:
            return None
        
        try:
            if hasattr(font_color, 'rgb'):
                return str(font_color.rgb) if font_color.rgb else None
            elif isinstance(font_color, str):
                return font_color
            else:
                return str(font_color) if font_color else None
        except (TypeError, ValueError):
            return None
    
    def map_formatting_to_translation(self, original_segments: List[FormattingSegment], 
                                     original_text: str, 
                                     translated_text: str) -> List[FormattingSegment]:
        """
        Сопоставляет форматирование оригинального текста с переведенным
        
        Args:
            original_segments: Сегменты форматирования оригинального текста
            original_text: Оригинальный текст
            translated_text: Переведенный текст
            
        Returns:
            Список сегментов форматирования для переведенного текста
        """
        if not original_segments:
            return [FormattingSegment(
                text=translated_text,
                start_pos=0,
                end_pos=len(translated_text)
            )]
        
        # Если у нас только один сегмент, применяем его форматирование ко всему переводу
        if len(original_segments) == 1:
            segment = original_segments[0]
            return [FormattingSegment(
                text=translated_text,
                bold=segment.bold,
                italic=segment.italic,
                underline=segment.underline,
                font_name=segment.font_name,
                font_size=segment.font_size,
                font_color=segment.font_color,
                start_pos=0,
                end_pos=len(translated_text)
            )]
        
        # Для множественных сегментов используем пропорциональное распределение
        return self._proportional_formatting_mapping(original_segments, original_text, translated_text)
    
    def _proportional_formatting_mapping(self, original_segments: List[FormattingSegment], 
                                       original_text: str, 
                                       translated_text: str) -> List[FormattingSegment]:
        """
        Пропорциональное сопоставление форматирования
        """
        if not translated_text.strip():
            return []
        
        total_original_length = len(original_text)
        total_translated_length = len(translated_text)
        
        if total_original_length == 0:
            return [FormattingSegment(text=translated_text, start_pos=0, end_pos=total_translated_length)]
        
        translated_segments = []
        current_translated_pos = 0
        
        for segment in original_segments:
            if not segment.text.strip():
                continue
            
            # Вычисляем пропорцию этого сегмента
            segment_proportion = len(segment.text) / total_original_length
            
            # Вычисляем длину соответствующего сегмента в переводе
            translated_segment_length = max(1, int(segment_proportion * total_translated_length))
            
            # Убеждаемся, что не выходим за границы
            if current_translated_pos + translated_segment_length > total_translated_length:
                translated_segment_length = total_translated_length - current_translated_pos
            
            if translated_segment_length > 0:
                segment_text = translated_text[current_translated_pos:current_translated_pos + translated_segment_length]
                
                translated_segment = FormattingSegment(
                    text=segment_text,
                    bold=segment.bold,
                    italic=segment.italic,
                    underline=segment.underline,
                    font_name=segment.font_name,
                    font_size=segment.font_size,
                    font_color=segment.font_color,
                    start_pos=current_translated_pos,
                    end_pos=current_translated_pos + translated_segment_length
                )
                
                translated_segments.append(translated_segment)
                current_translated_pos += translated_segment_length
        
        # Убеждаемся, что весь переведенный текст покрыт
        if current_translated_pos < total_translated_length:
            remaining_text = translated_text[current_translated_pos:]
            
            # Используем форматирование последнего сегмента
            last_segment = original_segments[-1] if original_segments else FormattingSegment(text="")
            
            remaining_segment = FormattingSegment(
                text=remaining_text,
                bold=last_segment.bold,
                italic=last_segment.italic,
                underline=last_segment.underline,
                font_name=last_segment.font_name,
                font_size=last_segment.font_size,
                font_color=last_segment.font_color,
                start_pos=current_translated_pos,
                end_pos=total_translated_length
            )
            
            translated_segments.append(remaining_segment)
        
        return translated_segments
    
    def apply_formatting_to_paragraph(self, paragraph: Paragraph, 
                                    formatted_segments: List[FormattingSegment],
                                    paragraph_alignment=None) -> bool:
        """
        Применяет форматирование к параграфу
        
        Args:
            paragraph: Параграф для форматирования
            formatted_segments: Сегменты с форматированием
            paragraph_alignment: Выравнивание параграфа
            
        Returns:
            True если форматирование применено успешно
        """
        try:
            # Очищаем параграф
            paragraph.clear()
            
            # Применяем выравнивание параграфа
            if paragraph_alignment is not None:
                paragraph.alignment = paragraph_alignment
            
            # Добавляем сегменты с форматированием
            for segment in formatted_segments:
                if not segment.text:
                    continue
                    
                run = paragraph.add_run(segment.text)
                
                # Применяем форматирование к run
                self._apply_run_formatting(run, segment)
            
            return True
            
        except Exception as e:
            print(f"Ошибка применения форматирования: {e}")
            return False
    
    def _apply_run_formatting(self, run: Run, segment: FormattingSegment):
        """Применяет форматирование к отдельному run - СТРОГИЙ КОНСЕРВАТИВНЫЙ режим"""
        try:
            # Применяем только базовое форматирование: жирный, курсив, подчеркивание.
            if segment.bold is not None:
                run.bold = segment.bold
            
            if segment.italic is not None:
                run.italic = segment.italic
            
            if segment.underline is not None:
                run.underline = segment.underline
            
            # --- ЦВЕТ И ШРИФТЫ ПОЛНОСТЬЮ ИГНОРИРУЮТСЯ ---
            # Это предотвращает "протекание" синего цвета и других стилей.
            
        except Exception as e:
            print(f"Ошибка применения базового форматирования к run: {e}")
    
    def _apply_font_color(self, run: Run, color_value: str):
        """Применяет цвет шрифта к run - ОТКЛЮЧЕНО для избежания синего выделения"""
        # НЕ применяем цвет для избежания синего выделения и других проблем
        pass
    
    def map_conservative_formatting_to_translation(self, original_segments: List[FormattingSegment], 
                                                  original_text: str, 
                                                  translated_text: str) -> List[FormattingSegment]:
        """
        КОНСЕРВАТИВНОЕ сопоставление форматирования - исправляет проблемы с синим выделением
        """
        if not original_segments or not translated_text.strip():
            return [FormattingSegment(
                text=translated_text,
                start_pos=0,
                end_pos=len(translated_text)
            )]
        
        # Если у оригинала сложное форматирование, используем упрощенный подход
        if len(original_segments) > 3:
            # Берем самый распространенный стиль
            most_common_style = self._get_most_common_style(original_segments)
            return [FormattingSegment(
                text=translated_text,
                bold=most_common_style.get('bold'),
                italic=most_common_style.get('italic'),
                underline=most_common_style.get('underline'),
                font_name=most_common_style.get('font_name'),
                font_size=most_common_style.get('font_size'),
                font_color=None,  # Убираем цвет, чтобы избежать синего выделения
                start_pos=0,
                end_pos=len(translated_text)
            )]
        
        # Для простого форматирования используем первый сегмент без цветов
        base_segment = original_segments[0]
        return [FormattingSegment(
            text=translated_text,
            bold=base_segment.bold,
            italic=base_segment.italic,
            underline=base_segment.underline,
            font_name=base_segment.font_name,
            font_size=base_segment.font_size,
            font_color=None,  # Убираем цвет
            start_pos=0,
            end_pos=len(translated_text)
        )]

    def _get_most_common_style(self, segments: List[FormattingSegment]) -> Dict[str, Any]:
        """Находит наиболее распространенный стиль среди сегментов"""
        styles = {}
        
        for segment in segments:
            style_key = f"{segment.bold}_{segment.italic}_{segment.underline}_{segment.font_name}_{segment.font_size}"
            if style_key not in styles:
                styles[style_key] = {
                    'count': 0,
                    'bold': segment.bold,
                    'italic': segment.italic,
                    'underline': segment.underline,
                    'font_name': segment.font_name,
                    'font_size': segment.font_size
                }
            styles[style_key]['count'] += 1
        
        # Возвращаем самый популярный стиль
        most_common = max(styles.values(), key=lambda x: x['count'])
        return most_common
    
    def analyze_formatting_complexity(self, formatting_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Анализирует сложность форматирования в документе
        
        Args:
            formatting_data: Данные форматирования
            
        Returns:
            Статистика форматирования
        """
        if not formatting_data or 'runs' not in formatting_data:
            return {
                'complexity': 'simple',
                'total_runs': 0,
                'has_bold': False,
                'has_italic': False,
                'has_underline': False,
                'unique_fonts': 0,
                'unique_sizes': 0,
                'unique_colors': 0
            }
        
        runs = formatting_data['runs']
        
        bold_count = sum(1 for run in runs if run.get('bold'))
        italic_count = sum(1 for run in runs if run.get('italic'))
        underline_count = sum(1 for run in runs if run.get('underline'))
        
        unique_fonts = len(set(run.get('font_name') for run in runs if run.get('font_name')))
        unique_sizes = len(set(str(run.get('font_size')) for run in runs if run.get('font_size')))
        unique_colors = len(set(run.get('font_color') for run in runs if run.get('font_color')))
        
        # Определяем сложность
        complexity = 'simple'
        if len(runs) > 3 or unique_fonts > 1 or unique_sizes > 1 or unique_colors > 1:
            complexity = 'medium'
        if len(runs) > 6 or unique_fonts > 2 or unique_sizes > 2 or unique_colors > 2:
            complexity = 'complex'
        
        return {
            'complexity': complexity,
            'total_runs': len(runs),
            'has_bold': bold_count > 0,
            'has_italic': italic_count > 0,
            'has_underline': underline_count > 0,
            'unique_fonts': unique_fonts,
            'unique_sizes': unique_sizes,
            'unique_colors': unique_colors,
            'bold_percentage': (bold_count / len(runs)) * 100 if runs else 0,
            'italic_percentage': (italic_count / len(runs)) * 100 if runs else 0,
            'underline_percentage': (underline_count / len(runs)) * 100 if runs else 0
        }
    
    def create_formatting_summary(self, all_elements_formatting: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Создает сводку форматирования для всего документа
        
        Args:
            all_elements_formatting: Список всех данных форматирования элементов
            
        Returns:
            Сводная статистика форматирования документа
        """
        total_elements = len(all_elements_formatting)
        
        if total_elements == 0:
            return {'total_elements': 0, 'formatting_complexity': 'none'}
        
        complexity_counts = {'simple': 0, 'medium': 0, 'complex': 0}
        total_runs = 0
        total_bold = 0
        total_italic = 0
        total_underline = 0
        all_fonts = set()
        all_sizes = set()
        all_colors = set()
        
        for formatting_data in all_elements_formatting:
            analysis = self.analyze_formatting_complexity(formatting_data)
            
            complexity_counts[analysis['complexity']] += 1
            total_runs += analysis['total_runs']
            
            if analysis['has_bold']:
                total_bold += 1
            if analysis['has_italic']:
                total_italic += 1
            if analysis['has_underline']:
                total_underline += 1
            
            # Собираем уникальные атрибуты форматирования
            if formatting_data and 'runs' in formatting_data:
                for run in formatting_data['runs']:
                    if run.get('font_name'):
                        all_fonts.add(run.get('font_name'))
                    if run.get('font_size'):
                        all_sizes.add(str(run.get('font_size')))
                    if run.get('font_color'):
                        all_colors.add(run.get('font_color'))
        
        # Определяем общую сложность документа
        if complexity_counts['complex'] > total_elements * 0.3:
            overall_complexity = 'complex'
        elif complexity_counts['medium'] > total_elements * 0.5:
            overall_complexity = 'medium'
        else:
            overall_complexity = 'simple'
        
        return {
            'total_elements': total_elements,
            'overall_complexity': overall_complexity,
            'complexity_distribution': complexity_counts,
            'total_runs': total_runs,
            'average_runs_per_element': total_runs / total_elements if total_elements > 0 else 0,
            'elements_with_bold': total_bold,
            'elements_with_italic': total_italic,
            'elements_with_underline': total_underline,
            'unique_fonts': len(all_fonts),
            'unique_font_sizes': len(all_sizes),
            'unique_colors': len(all_colors),
            'fonts_used': list(all_fonts),
            'bold_percentage': (total_bold / total_elements) * 100 if total_elements > 0 else 0,
                          'italic_percentage': (total_italic / total_elements) * 100 if total_elements > 0 else 0,
              'underline_percentage': (total_underline / total_elements) * 100 if total_elements > 0 else 0
         }    
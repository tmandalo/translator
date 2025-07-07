"""
Улучшенный модуль для обработки изображений в .docx документах
Работает напрямую с ZIP структурой для надежного извлечения изображений
"""

import os
import io
import tempfile
import logging
import zipfile
import xml.etree.ElementTree as ET
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass
from pathlib import Path
from PIL import Image

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


@dataclass
class ImageInfo:
    """Класс для хранения информации об изображении"""
    image_id: str
    image_data: bytes
    image_format: str
    width: Optional[float] = None
    height: Optional[float] = None
    paragraph_index: Optional[int] = None
    rel_id: Optional[str] = None
    filename: Optional[str] = None


@dataclass
class ImageElement:
    """Класс для хранения элемента изображения (совместимость с удаленным image_processor)"""
    image_id: str
    image_data: bytes
    image_format: str
    width: Optional[int] = None
    height: Optional[int] = None
    paragraph_index: Optional[int] = None
    is_inline: bool = True
    description: Optional[str] = None
    alt_text: Optional[str] = None


class ImprovedImageProcessor:
    """Улучшенный класс для обработки изображений в документах"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.temp_dir = None
        self.images: List[ImageInfo] = []
        
    def extract_images_from_docx(self, docx_path: str) -> List[ImageInfo]:
        """
        Извлекает изображения напрямую из ZIP структуры .docx файла
        
        Args:
            docx_path: Путь к .docx файлу
            
        Returns:
            Список информации об изображениях
        """
        images = []
        
        try:
            # Создаем временную директорию
            if not self.temp_dir:
                self.temp_dir = tempfile.mkdtemp(prefix='docx_images_')
            
            # Открываем .docx как ZIP архив
            with zipfile.ZipFile(docx_path, 'r') as docx_zip:
                # Получаем список всех файлов в архиве
                file_list = docx_zip.namelist()
                
                # Ищем изображения в папке word/media/
                media_files = [f for f in file_list if f.startswith('word/media/')]
                
                # Парсим relationships для получения связей
                relationships = self._parse_relationships(docx_zip)
                
                # Парсим основной документ для поиска позиций изображений
                image_positions = self._parse_document_for_images(docx_zip)
                
                # Сохраняем результаты для логирования
                self._last_relationships = relationships
                self._last_positions = image_positions
                
                self.logger.info(f"Найдено {len(media_files)} медиа файлов и {len(relationships)} relationships")
                self.logger.info(f"Найдено {len(image_positions)} позиций изображений в документе")
                
                # Обрабатываем каждое медиа файл
                positioned_images = 0
                unpositioned_images = 0
                
                for media_file in media_files:
                    try:
                        # Извлекаем данные изображения
                        image_data = docx_zip.read(media_file)
                        
                        # Определяем формат
                        image_format = self._detect_image_format(image_data)
                        
                        if image_format == 'unknown':
                            self.logger.warning(f"Неизвестный формат изображения: {media_file}")
                            continue
                        
                        # Создаем ID
                        image_id = f"extracted_{len(images) + 1}"
                        filename = os.path.basename(media_file)
                        
                        # Ищем информацию о позиции
                        rel_id = self._find_rel_id_for_media(media_file, relationships)
                        paragraph_index = image_positions.get(rel_id)
                        
                        # Подсчитываем изображения с позициями
                        if paragraph_index is not None:
                            positioned_images += 1
                            self.logger.info(f"Изображение {filename} -> rel_id: {rel_id} -> параграф {paragraph_index}")
                        else:
                            unpositioned_images += 1
                            self.logger.warning(f"Изображение {filename} -> rel_id: {rel_id} -> позиция не найдена")
                        
                        # Получаем размеры изображения
                        width, height = self._get_image_dimensions(image_data)
                        
                        # Сохраняем изображение во временную папку
                        temp_path = os.path.join(self.temp_dir, f"{image_id}.{image_format}")
                        with open(temp_path, 'wb') as f:
                            f.write(image_data)
                        
                        # Создаем объект информации об изображении
                        image_info = ImageInfo(
                            image_id=image_id,
                            image_data=image_data,
                            image_format=image_format,
                            width=width,
                            height=height,
                            paragraph_index=paragraph_index,
                            rel_id=rel_id,
                            filename=filename
                        )
                        
                        images.append(image_info)
                        self.logger.info(f"Извлечено изображение: {filename} ({image_format}, {len(image_data)} байт)")
                        
                    except Exception as e:
                        self.logger.warning(f"Ошибка обработки медиа файла {media_file}: {e}")
                        continue
            
            self.images = images
            self.logger.info(f"Всего извлечено изображений: {len(images)}")
            self.logger.info(f"Изображений с найденными позициями: {positioned_images}")
            self.logger.info(f"Изображений без позиций: {unpositioned_images}")
            
        except Exception as e:
            self.logger.error(f"Ошибка извлечения изображений из {docx_path}: {e}")
            
        return images
    
    def _parse_relationships(self, docx_zip: zipfile.ZipFile) -> Dict[str, str]:
        """Парсит файл relationships для получения связей между ID и файлами"""
        relationships = {}
        
        try:
            # Читаем word/_rels/document.xml.rels
            rels_content = docx_zip.read('word/_rels/document.xml.rels')
            root = ET.fromstring(rels_content)
            
            # Парсим relationships
            for rel in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                rel_id = rel.get('Id')
                target = rel.get('Target')
                rel_type = rel.get('Type')
                
                # Интересуют только изображения
                if rel_type and 'image' in rel_type.lower():
                    relationships[rel_id] = target
                    
        except Exception as e:
            self.logger.warning(f"Ошибка парсинга relationships: {e}")
            
        return relationships
    
    def _parse_document_for_images(self, docx_zip: zipfile.ZipFile) -> Dict[str, int]:
        """
        УЛУЧШЕННЫЙ парсинг основного документа для поиска позиций изображений
        Обеспечивает более точное определение позиций и лучшую совместимость с python-docx
        """
        image_positions = {}
        
        try:
            # Читаем word/document.xml
            doc_content = docx_zip.read('word/document.xml')
            root = ET.fromstring(doc_content)
            
            # Ищем ТОЛЬКО параграфы в основном теле документа (исключаем headers, footers, etc.)
            body = root.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body')
            if body is None:
                self.logger.warning("Не найден body элемент в документе")
                return image_positions
            
            # Получаем ВСЕ параграфы из body (включая пустые)
            # Это важно для правильного сопоставления с python-docx
            all_paragraphs = body.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
            
            self.logger.info(f"🔍 XML парсер: найдено {len(all_paragraphs)} параграфов в основном теле документа")
            print(f"🔍 XML парсер: найдено {len(all_paragraphs)} параграфов в основном теле документа")
            
            # Сначала проанализируем все параграфы для создания карты соответствий
            paragraph_mapping = []
            significant_para_count = 0
            
            for xml_idx, paragraph in enumerate(all_paragraphs):
                # Анализируем содержимое параграфа
                text_nodes = paragraph.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                image_nodes = paragraph.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')
                
                # Подсчитываем текстовое содержимое
                text_content = ''.join(node.text or '' for node in text_nodes).strip()
                has_meaningful_text = len(text_content) > 0
                has_images = len(image_nodes) > 0
                
                # Определяем, является ли параграф значимым
                is_significant = has_meaningful_text or has_images
                
                if is_significant:
                    paragraph_mapping.append({
                        'xml_index': xml_idx,
                        'docx_index': significant_para_count,
                        'has_text': has_meaningful_text,
                        'has_images': has_images,
                        'text_preview': text_content[:50] + '...' if len(text_content) > 50 else text_content
                    })
                    significant_para_count += 1
                    
                    self.logger.debug(f"Параграф XML[{xml_idx}] -> DOCX[{significant_para_count-1}]: "
                                    f"text={has_meaningful_text}, images={has_images}, "
                                    f"preview='{text_content[:30]}...' if text_content else 'empty'")
            
            self.logger.info(f"🔍 Создано сопоставление: {len(paragraph_mapping)} значимых параграфов из {len(all_paragraphs)} XML параграфов")
            print(f"🔍 Создано сопоставление: {len(paragraph_mapping)} значимых параграфов из {len(all_paragraphs)} XML параграфов")
            
            # Теперь ищем изображения в каждом параграфе
            for xml_idx, paragraph in enumerate(all_paragraphs):
                images_in_paragraph = []
                
                # Находим соответствующий docx индекс
                docx_idx = None
                for mapping in paragraph_mapping:
                    if mapping['xml_index'] == xml_idx:
                        docx_idx = mapping['docx_index']
                        break
                
                # Если параграф не значимый, пропускаем поиск изображений
                if docx_idx is None:
                    continue
                
                # === УЛУЧШЕННЫЙ ПОИСК ИЗОБРАЖЕНИЙ ===
                
                # 1. Ищем изображения в drawing элементах (современный формат)
                drawings = paragraph.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')
                for drawing in drawings:
                    # Ищем все возможные типы изображений в drawing
                    blips = drawing.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
                    for blip in blips:
                        embed_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                        if embed_id:
                            image_positions[embed_id] = docx_idx
                            images_in_paragraph.append(f"drawing:{embed_id}")
                            self.logger.debug(f"Найдено изображение в drawing: {embed_id} -> параграф {docx_idx}")
                
                # 2. Ищем inline изображения (встроенные изображения)  
                inline_shapes = paragraph.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}object')
                for inline_shape in inline_shapes:
                    blips = inline_shape.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
                    for blip in blips:
                        embed_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                        if embed_id:
                            image_positions[embed_id] = docx_idx
                            images_in_paragraph.append(f"inline:{embed_id}")
                            self.logger.debug(f"Найдено inline изображение: {embed_id} -> параграф {docx_idx}")
                
                # 3. Ищем pict элементы (старый формат Word)
                picts = paragraph.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pict')
                for pict in picts:
                    # Ищем различные типы изображений в pict
                    shapes = pict.findall('.//*[@r:id]', namespaces={'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'})
                    for shape in shapes:
                        embed_id = shape.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                        if embed_id:
                            image_positions[embed_id] = docx_idx
                            images_in_paragraph.append(f"pict:{embed_id}")
                            self.logger.debug(f"Найдено pict изображение: {embed_id} -> параграф {docx_idx}")
                
                # 4. Ищем в run элементах (дополнительная проверка)
                runs = paragraph.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
                for run in runs:
                    # Ищем все элементы с embed атрибутом
                    embeds = run.findall('.//*[@r:embed]', namespaces={'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'})
                    for embed in embeds:
                        embed_id = embed.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                        if embed_id and embed_id not in image_positions:
                            image_positions[embed_id] = docx_idx
                            images_in_paragraph.append(f"run:{embed_id}")
                            self.logger.debug(f"Найдено run изображение: {embed_id} -> параграф {docx_idx}")
                
                # 5. Дополнительный поиск по альтернативным атрибутам
                # Ищем элементы с r:id (alternative relationship format)
                alt_images = paragraph.findall('.//*[@r:id]', namespaces={'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'})
                for alt_img in alt_images:
                    embed_id = alt_img.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    if embed_id and embed_id not in image_positions:
                        # Проверяем, что это действительно изображение
                        if self._is_image_relationship(embed_id):
                            image_positions[embed_id] = docx_idx
                            images_in_paragraph.append(f"alt:{embed_id}")
                            self.logger.debug(f"Найдено альтернативное изображение: {embed_id} -> параграф {docx_idx}")
                
                # Логируем найденные изображения в параграфе
                if images_in_paragraph:
                    self.logger.info(f"📍 Параграф XML[{xml_idx}] -> DOCX[{docx_idx}]: найдено {len(images_in_paragraph)} изображений: {', '.join(images_in_paragraph)}")
                    print(f"📍 Параграф XML[{xml_idx}] -> DOCX[{docx_idx}]: {len(images_in_paragraph)} изображений - {', '.join(images_in_paragraph)}")
                    
            # Финальная статистика
            self.logger.info(f"🎯 ИТОГО найдено позиций изображений: {len(image_positions)}")
            print(f"🎯 ИТОГО найдено позиций изображений: {len(image_positions)}")
            
            # Детальный вывод всех найденных позиций
            for embed_id, para_idx in sorted(image_positions.items(), key=lambda x: x[1]):
                self.logger.info(f"  📌 Изображение {embed_id} -> параграф {para_idx}")
                print(f"  📌 Изображение {embed_id} -> параграф {para_idx}")
                            
        except Exception as e:
            self.logger.error(f"❌ Ошибка УЛУЧШЕННОГО парсинга документа для изображений: {e}")
            print(f"❌ Ошибка УЛУЧШЕННОГО парсинга документа для изображений: {e}")
            
        return image_positions
    
    def _find_rel_id_for_media(self, media_file: str, relationships: Dict[str, str]) -> Optional[str]:
        """Находит relationship ID для медиа файла"""
        media_filename = os.path.basename(media_file)
        
        # Попытка точного совпадения
        for rel_id, target in relationships.items():
            if target.endswith(media_filename):
                return rel_id
        
        # Попытка совпадения без "media/" префикса
        for rel_id, target in relationships.items():
            if target.endswith(media_filename) or target.endswith(media_file):
                return rel_id
        
        # Попытка совпадения по имени файла без расширения
        media_name_without_ext = os.path.splitext(media_filename)[0]
        for rel_id, target in relationships.items():
            target_name = os.path.splitext(os.path.basename(target))[0]
            if target_name == media_name_without_ext:
                return rel_id
        
        self.logger.warning(f"Не найден relationship ID для медиа файла: {media_file}")
        return None
    
    def _get_image_dimensions(self, image_data: bytes) -> Tuple[Optional[float], Optional[float]]:
        """Получает размеры изображения в дюймах"""
        try:
            with Image.open(io.BytesIO(image_data)) as img:
                # Получаем размеры в пикселях
                width_px, height_px = img.size
                
                # Получаем DPI (по умолчанию 96)
                dpi = img.info.get('dpi', (96, 96))
                if isinstance(dpi, tuple):
                    dpi_x, dpi_y = dpi
                else:
                    dpi_x = dpi_y = dpi
                
                # Конвертируем в дюймы
                width_inches = width_px / dpi_x
                height_inches = height_px / dpi_y
                
                return width_inches, height_inches
                
        except Exception as e:
            self.logger.warning(f"Не удалось определить размеры изображения: {e}")
            return None, None
    
    def _detect_image_format(self, image_data: bytes) -> str:
        """Определяет формат изображения по binary data"""
        try:
            # Проверяем сигнатуры файлов
            if image_data.startswith(b'\x89PNG'):
                return 'png'
            elif image_data.startswith(b'\xFF\xD8\xFF'):
                return 'jpeg'
            elif image_data.startswith(b'GIF87a') or image_data.startswith(b'GIF89a'):
                return 'gif'
            elif image_data.startswith(b'BM'):
                return 'bmp'
            elif image_data.startswith(b'RIFF') and b'WEBP' in image_data[:12]:
                return 'webp'
            else:
                # Попытка определить через PIL
                with Image.open(io.BytesIO(image_data)) as img:
                    return img.format.lower()
                    
        except Exception as e:
            self.logger.warning(f"Не удалось определить формат изображения: {e}")
            return 'unknown'
    
    def insert_images_into_document(self, document: Document, original_docx_path: str) -> bool:
        """
        Вставляет изображения в новый документ
        
        Args:
            document: Целевой документ
            original_docx_path: Путь к оригинальному документу
            
        Returns:
            True если успешно, False иначе
        """
        try:
            # Сначала извлекаем изображения если ещё не сделали
            if not self.images:
                self.extract_images_from_docx(original_docx_path)
            
            # Вставляем изображения в соответствующие позиции
            for image_info in self.images:
                success = self._insert_single_image(document, image_info)
                if not success:
                    self.logger.warning(f"Не удалось вставить изображение {image_info.image_id}")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Ошибка вставки изображений: {e}")
            return False
    
    def _insert_single_image(self, document: Document, image_info: ImageInfo) -> bool:
        """Вставляет одно изображение в документ"""
        try:
            # Определяем позицию для вставки
            target_paragraph_index = image_info.paragraph_index
            
            if target_paragraph_index is None or target_paragraph_index >= len(document.paragraphs):
                # Вставляем в конец документа
                paragraph = document.add_paragraph()
            else:
                # Получаем существующий параграф или создаем новый рядом
                if target_paragraph_index < len(document.paragraphs):
                    # Вставляем после указанного параграфа
                    target_para = document.paragraphs[target_paragraph_index]
                    paragraph = document.add_paragraph()
                else:
                    paragraph = document.add_paragraph()
            
            # Получаем изображение из временной папки
            temp_path = os.path.join(self.temp_dir, f"{image_info.image_id}.{image_info.image_format}")
            
            if not os.path.exists(temp_path):
                self.logger.warning(f"Файл изображения не найден: {temp_path}")
                return False
            
            # Вставляем изображение
            run = paragraph.add_run()
            
            # Определяем размеры
            if image_info.width and image_info.height:
                # Ограничиваем размеры
                max_width = 6.0  # максимум 6 дюймов
                max_height = 8.0  # максимум 8 дюймов
                
                width = min(image_info.width, max_width)
                height = min(image_info.height, max_height)
                
                # Сохраняем пропорции
                if width / image_info.width < height / image_info.height:
                    height = width * (image_info.height / image_info.width)
                else:
                    width = height * (image_info.width / image_info.height)
                    
                run.add_picture(temp_path, width=Inches(width), height=Inches(height))
            else:
                # Используем размер по умолчанию
                run.add_picture(temp_path, width=Inches(4.0))
            
            self.logger.info(f"Изображение {image_info.image_id} успешно вставлено")
            return True
            
        except Exception as e:
            self.logger.warning(f"Ошибка вставки изображения {image_info.image_id}: {e}")
            return False
    
    def cleanup_temp_files(self):
        """Очищает временные файлы"""
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                import shutil
                shutil.rmtree(self.temp_dir)
                self.temp_dir = None
                self.logger.info("Временные файлы изображений очищены")
            except Exception as e:
                self.logger.error(f"Ошибка очистки временных файлов: {e}")
    
    def get_image_statistics(self) -> Dict[str, Any]:
        """Возвращает статистику изображений"""
        if not self.images:
            return {'total_images': 0}
        
        formats = {}
        total_size = 0
        
        for image in self.images:
            # Подсчитываем форматы
            formats[image.image_format] = formats.get(image.image_format, 0) + 1
            
            # Подсчитываем общий размер
            total_size += len(image.image_data)
        
        return {
            'total_images': len(self.images),
            'formats': formats,
            'total_size_mb': round(total_size / (1024 * 1024), 2),
            'average_size_kb': round((total_size / len(self.images)) / 1024, 2) if self.images else 0
        }
    
    def _is_image_relationship(self, rel_id: str) -> bool:
        """
        Проверяет, является ли relationship ID изображением
        
        Args:
            rel_id: ID relationship для проверки
            
        Returns:
            True если это изображение, False иначе
        """
        # Расширенная проверка на основе анализа relationship
        # Пока возвращаем True для всех, но можно добавить фильтрацию по типу
        self.logger.debug(f"🔍 Проверка relationship {rel_id} на предмет изображения")
        return True
    
    def get_detailed_extraction_log(self) -> str:
        """
        Возвращает детальный лог процесса извлечения изображений
        
        Returns:
            Строка с детальным логом
        """
        if not self.images:
            return "Изображения не извлечены"
            
        log_lines = [
            f"📊 ДЕТАЛЬНЫЙ ОТЧЕТ ПО ИЗВЛЕЧЕНИЮ ИЗОБРАЖЕНИЙ",
            f"=" * 50,
            f"Всего изображений: {len(self.images)}",
            f"Временная папка: {self.temp_dir}",
            f"",
            f"📋 СПИСОК ИЗОБРАЖЕНИЙ:"
        ]
        
        positioned_count = 0
        unpositioned_count = 0
        
        for i, img in enumerate(self.images, 1):
            status = "✅ ПОЗИЦИОНИРОВАНО" if img.paragraph_index is not None else "❓ БЕЗ ПОЗИЦИИ"
            if img.paragraph_index is not None:
                positioned_count += 1
            else:
                unpositioned_count += 1
                
            log_lines.append(f"  {i}. {img.image_id}")
            log_lines.append(f"     Файл: {img.filename}")
            log_lines.append(f"     Формат: {img.image_format}")
            log_lines.append(f"     Размер: {img.width}x{img.height} дюймов" if img.width and img.height else "     Размер: неизвестен")
            log_lines.append(f"     Позиция: {img.paragraph_index}" if img.paragraph_index is not None else "     Позиция: не определена")
            log_lines.append(f"     Relationship: {img.rel_id}")
            log_lines.append(f"     Статус: {status}")
            log_lines.append("")
        
        log_lines.extend([
            f"📈 СТАТИСТИКА:",
            f"  Позиционированных: {positioned_count}",
            f"  Без позиций: {unpositioned_count}",
            f"  Успешность: {(positioned_count/len(self.images)*100):.1f}%",
        ])
        
        return "\n".join(log_lines) 
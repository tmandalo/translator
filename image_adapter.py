"""
Адаптер для преобразования ImageInfo в ImageElement
"""

from typing import List, Optional
from improved_image_processor import ImageInfo, ImageElement


class ImageAdapter:
    """Адаптер для преобразования между ImageInfo и ImageElement"""
    
    @staticmethod
    def convert_to_image_element(image_info: ImageInfo) -> ImageElement:
        """
        Преобразует ImageInfo в ImageElement для совместимости
        
        Args:
            image_info: Информация об изображении из улучшенного процессора
            
        Returns:
            Элемент изображения в старом формате
        """
        # Логируем процесс преобразования
        print(f"🔄 ADAPTER: Преобразование {image_info.image_id} (позиция: {image_info.paragraph_index})")
        
        return ImageElement(
            image_id=image_info.image_id,
            image_data=image_info.image_data,
            image_format=image_info.image_format,
            width=int(image_info.width * 96) if image_info.width else None,  # Конвертируем дюймы в пиксели
            height=int(image_info.height * 96) if image_info.height else None,  # Конвертируем дюймы в пиксели
            paragraph_index=image_info.paragraph_index,
            is_inline=True,
            description=image_info.filename,
            alt_text=f"Изображение {image_info.filename}"
        )
    
    @staticmethod
    def convert_list_to_image_elements(image_infos: List[ImageInfo]) -> List[ImageElement]:
        """
        Преобразует список ImageInfo в список ImageElement
        
        Args:
            image_infos: Список информации об изображениях
            
        Returns:
            Список элементов изображений в старом формате
        """
        print(f"🔄 ADAPTER: Преобразование {len(image_infos)} изображений из ImageInfo в ImageElement")
        
        elements = [ImageAdapter.convert_to_image_element(info) for info in image_infos]
        
        # Статистика преобразования
        positioned_count = len([elem for elem in elements if elem.paragraph_index is not None])
        unpositioned_count = len(elements) - positioned_count
        
        print(f"🔄 ADAPTER: Результат - {positioned_count} с позициями, {unpositioned_count} без позиций")
        
        return elements 
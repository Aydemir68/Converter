#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import argparse
import glob
from pathlib import Path
import time
import string
import re

# Import conversion functions from main script
from Script import (
    ocr_pdf_to_txt, 
    convert_pdf_to_txt_direct, 
    convert_pdf_to_docx_then_txt,
    extract_text_from_pdf_pypdf,
    convert_docx_to_txt
)

def batch_convert(input_folder, output_folder, method='auto', pattern='*.pdf'):
    """
    Batch convert PDF or DOCX files to TXT
    Args:
        input_folder: Folder containing files
        output_folder: Folder to save TXT files
        method: Conversion method ('auto', 'direct', 'ocr', 'docx', 'docx2txt')
        pattern: File pattern to match (default: *.pdf or *.docx)
    """
    
    # Find all files by pattern
    files = glob.glob(os.path.join(input_folder, pattern))
    files.extend(glob.glob(os.path.join(input_folder, pattern.upper())))
    
    if not files:
        print(f"❌ Файлы не найдены в папке: {input_folder}")
        return
    
    print(f"📁 Найдено {len(files)} файлов по шаблону {pattern}")
    print(f"📂 Папка ввода: {input_folder}")
    print(f"📂 Папка вывода: {output_folder}")
    print(f"🔧 Метод конвертации: {method}")
    print("-" * 60)
    
    # Create output folder
    os.makedirs(output_folder, exist_ok=True)
    
    successful_conversions = []
    failed_conversions = []
    
    for i, file_path in enumerate(files, 1):
        filename = os.path.basename(file_path)
        print(f"[{i}/{len(files)}] Обрабатывается: {filename}")
        txt_path = os.path.join(output_folder, os.path.splitext(filename)[0] + '.txt')
        try:
            ext = os.path.splitext(filename)[1].lower()
            if ext == '.pdf':
                # Determine conversion method if auto
                if method == 'auto':
                    # Try direct extraction first
                    try:
                        text = extract_text_from_pdf_pypdf(file_path)
                        if len(text.strip()) > 50:  # If we got substantial text
                            conversion_method = 'direct'
                        else:
                            conversion_method = 'ocr'
                    except:
                        conversion_method = 'ocr'
                else:
                    conversion_method = method
                
                # Perform conversion
                if conversion_method == 'direct':
                    success, message = convert_pdf_to_txt_direct(file_path, output_folder)
                elif conversion_method == 'ocr':
                    success, message = ocr_pdf_to_txt(file_path, output_folder)
                elif conversion_method == 'docx':
                    success, message = convert_pdf_to_docx_then_txt(file_path, output_folder)
                else:
                    raise Exception(f"Неизвестный метод конвертации: {conversion_method}")
            elif ext == '.docx':
                success, message = convert_docx_to_txt(file_path, output_folder)
            else:
                raise Exception(f"Неизвестный тип файла: {filename}")
            
            # Проверка на пустой результат и мусор
            is_empty = False
            is_garbage = False
            if not os.path.exists(txt_path) or os.path.getsize(txt_path) == 0:
                is_empty = True
            else:
                with open(txt_path, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                    # Удаляем все невидимые символы (пробелы, табуляции, переносы строк, неразрывные пробелы, zero-width и т.д.)
                    content_no_invisible = re.sub(r'[\s\u00A0\u200B\u200C\u200D\uFEFF]', '', content)
                    # Удаляем все неотображаемые символы (ord < 32, кроме \n, \r, \t)
                    content_no_control = ''.join(c for c in content_no_invisible if ord(c) >= 32 or c in '\n\r\t')
                    # Оставляем только буквы и цифры (латиница, кириллица, цифры)
                    letters_digits = re.findall(r'[A-Za-zА-Яа-яЁё0-9]', content_no_control)
                    num_letters_digits = len(letters_digits)
                    total_chars = len(content)
                    # Пустой, если букв/цифр < 10
                    if num_letters_digits < 10:
                        is_empty = True
                    else:
                        # Мусор, если доля букв/цифр < 0.3
                        if total_chars > 0 and (num_letters_digits / total_chars) < 0.3:
                            is_garbage = True
                        # Дополнительно: если есть длинные последовательности спецсимволов и нет ни кириллицы, ни латиницы
                        has_garbage_seq = re.search(r'[^\wА-Яа-яЁё]{4,}', content)
                        has_letters = re.search(r'[A-Za-zА-Яа-яЁё]', content)
                        if has_garbage_seq and not has_letters:
                            is_garbage = True
            
            if is_empty:
                failed_conversions.append(filename)
                print(f"   ❌ Ошибка: файл сконвертирован пустым!")
                if os.path.exists(txt_path):
                    os.remove(txt_path)
            elif is_garbage:
                failed_conversions.append(filename)
                print(f"   ❌ Ошибка: файл содержит мусор (неотображаемые символы или набор спецсимволов)!")
                if os.path.exists(txt_path):
                    os.remove(txt_path)
            elif success:
                successful_conversions.append(filename)
                print(f"   ✅ Успешно: {message}")
            else:
                failed_conversions.append(filename)
                print(f"   ❌ Ошибка: {message}")
        except Exception as e:
            failed_conversions.append(filename)
            print(f"   ❌ Ошибка: {str(e)}")
    
    # Print summary
    print("\n" + "=" * 60)
    print("📊 РЕЗУЛЬТАТЫ КОНВЕРТАЦИИ")
    print("=" * 60)
    print(f"✅ Успешно конвертировано: {len(successful_conversions)}")
    print(f"❌ Ошибок: {len(failed_conversions)}")
    
    if failed_conversions:
        print(f"\n📋 Файлы с ошибками:")
        for failed_file in failed_conversions:
            print(f"   • {failed_file}")
        
        # Save error report
        error_report_path = os.path.join(output_folder, "error_report.txt")
        with open(error_report_path, 'w', encoding='utf-8') as f:
            f.write("Отчет об ошибках конвертации\n")
            f.write("=" * 40 + "\n\n")
            f.write(f"Дата: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Папка ввода: {input_folder}\n")
            f.write(f"Папка вывода: {output_folder}\n")
            f.write(f"Метод конвертации: {method}\n")
            f.write(f"Всего файлов: {len(files)}\n")
            f.write(f"Успешно: {len(successful_conversions)}\n")
            f.write(f"Ошибок: {len(failed_conversions)}\n\n")
            f.write("Список файлов с ошибками (копируйте для поиска):\n")
            for failed_file in failed_conversions:
                fail_path = os.path.join(output_folder, os.path.splitext(failed_file)[0] + '.txt')
                reason = ''
                if not os.path.exists(fail_path) or os.path.getsize(fail_path) == 0:
                    reason = ' (пустой)'
                else:
                    with open(fail_path, 'r', encoding='utf-8', errors='ignore') as ftxt:
                        txt_content = ftxt.read()
                        txt_content_no_invisible = re.sub(r'[\s\u00A0\u200B\u200C\u200D\uFEFF]', '', txt_content)
                        txt_content_no_control = ''.join(c for c in txt_content_no_invisible if ord(c) >= 32 or c in '\n\r\t')
                        if txt_content_no_control == '':
                            reason = ' (пустой)'
                        else:
                            total_chars = len(txt_content)
                            invisible_count = sum(1 for c in txt_content if (ord(c) < 32 and c not in '\n\r\t'))
                            if total_chars > 0 and invisible_count / total_chars > 0.5:
                                reason = ' (мусор)'
                            has_cyrillic = re.search(r'[а-яА-ЯёЁ]', txt_content)
                            has_garbage_seq = re.search(r'[^\wа-яА-ЯёЁ]{2,}', txt_content)
                            if has_garbage_seq and not has_cyrillic:
                                reason = ' (мусор)'
                f.write(f"{failed_file}{reason}\n")
        
        print(f"\n📄 Отчет об ошибках сохранен в: {error_report_path}")
    
    if successful_conversions:
        print(f"\n✅ Успешно конвертированные файлы:")
        for success_file in successful_conversions:
            print(f"   • {success_file}")

def main():
    parser = argparse.ArgumentParser(
        description='Пакетный конвертер PDF в TXT',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры использования:
  python batch_converter.py /path/to/pdfs /path/to/output
  python batch_converter.py /path/to/pdfs /path/to/output --method ocr
  python batch_converter.py /path/to/pdfs /path/to/output --method direct --pattern "*.PDF"
        """
    )
    
    parser.add_argument('input_folder', help='Папка с PDF файлами')
    parser.add_argument('output_folder', help='Папка для сохранения TXT файлов')
    parser.add_argument('--method', choices=['auto', 'direct', 'ocr', 'docx', 'docx2txt'], 
                       default='auto', help='Метод конвертации (по умолчанию: auto)')
    parser.add_argument('--pattern', default='*.pdf', 
                       help='Шаблон файлов (по умолчанию: *.pdf или *.docx)')
    
    args = parser.parse_args()
    
    # Validate input folder
    if not os.path.exists(args.input_folder):
        print(f"❌ Ошибка: Папка ввода не существует: {args.input_folder}")
        sys.exit(1)
    
    if not os.path.isdir(args.input_folder):
        print(f"❌ Ошибка: Путь не является папкой: {args.input_folder}")
        sys.exit(1)
    
    # Start conversion
    try:
        batch_convert(args.input_folder, args.output_folder, args.method, args.pattern)
    except KeyboardInterrupt:
        print("\n⚠️  Конвертация прервана пользователем")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Критическая ошибка: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main() 
import os
import threading
import ctypes
import flet as ft
import xlwings as xw
import sys
import winsound
import re
import time

try:
    myappid = 'company.ezsearch.checker.v1' 
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception:
    pass

def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def exact_match(term, name):
    pattern = r"(^|[^a-z0-9])" + re.escape(term) + r"([^a-z0-9]|$)"
    return re.search(pattern, name) is not None

def main(page: ft.Page):
    page.title = "EZ Search V6"
    page.window.icon = "assets/icon.ico" 
    page.window.width = 650
    page.window.height = 700
    page.window.min_width = 500
    page.window.min_height = 600
    page.theme_mode = ft.ThemeMode.DARK 
    page.bgcolor = "#06090e" 
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.padding = 30

    current_lang = "AR"

    texts = {
        "AR": {
            "textfield_label": "قم بوضع أسماء الريفرنسات هنا (كل اسم في سطر)",
            "btn_search": "ابحث في الملفات واستخرج الإكسيل",
            "btn_stop": "إيقاف البحث",
            "status_searching": "جاري فحص {} ريفرنس... برجاء الانتظار",
            "snack_empty": "برجاء كتابة أو لصق الريفرنسات أولاً!",
            "snack_no_drive": "لم يتم العثور على مسار OneDrive الخاص بالشركة على هذا الجهاز.",
            "snack_success": "تم فحص {} ريفرنس واستخراج الإكسيل بنجاح!",
            "snack_error": "حدث خطأ أثناء الفحص: {}",
            "snack_stopped": "تم إيقاف البحث بناءً على طلبك!",
            "lang_btn_text": "English"
        },
        "EN": {
            "textfield_label": "Paste your references here (one name per line)",
            "btn_search": "Search Files & Generate Excel",
            "btn_stop": "Stop Search",
            "status_searching": "Searching {} references... Please wait",
            "snack_empty": "Please type or paste references first!",
            "snack_no_drive": "OneDrive path not found on this device.",
            "snack_success": "Successfully checked {} references and generated Excel!",
            "snack_error": "An error occurred during scan: {}",
            "snack_stopped": "Search stopped by user!",
            "lang_btn_text": "عربي"
        }
    }

    stop_search_flag = False 

    def show_snack(message, is_error=False):
        snack = ft.SnackBar(
            content=ft.Text(message, color=ft.Colors.WHITE, weight=ft.FontWeight.BOLD),
            bgcolor=ft.Colors.RED_700 if is_error else ft.Colors.GREEN_700,
            duration=4000
        )
        page.overlay.append(snack)
        snack.open = True
        page.update()

    logo_image = ft.Image(src="icon.png", width=130, height=130, fit="contain")

    def clear_all_refs(e):
        txt_refs.value = ""
        txt_refs.focus()
        page.update()

    txt_refs = ft.TextField(
        label=texts[current_lang]["textfield_label"],
        multiline=True,
        min_lines=18,
        max_lines=18,
        border_color=ft.Colors.BLUE_500,
        border_width=2,
        border_radius=20,
        expand=True,
        text_align=ft.TextAlign.LEFT,
        suffix=ft.TextButton(content=ft.Text("❌", size=16), on_click=clear_all_refs, tooltip="مسح الكل / Clear All")
    )

    pr = ft.ProgressBar(width=400, visible=False, color=ft.Colors.BLUE_700, bgcolor=ft.Colors.BLUE_100)
    status_text = ft.Text("", visible=False, color=ft.Colors.GREEN_600, weight=ft.FontWeight.BOLD)

    def toggle_language(e):
        nonlocal current_lang
        current_lang = "EN" if current_lang == "AR" else "AR"
        txt_refs.label = texts[current_lang]["textfield_label"]
        btn_check.content.value = texts[current_lang]["btn_search"]
        lang_btn.content.value = texts[current_lang]["lang_btn_text"]
        btn_stop.content.value = texts[current_lang]["btn_stop"]
        page.update()

    def stop_search_click(e):
        nonlocal stop_search_flag
        stop_search_flag = True
        btn_stop.style = ft.ButtonStyle(color=ft.Colors.GREY_700, bgcolor=ft.Colors.TRANSPARENT, elevation=0)
        btn_stop.disabled = True 
        page.update()

    def toggle_theme(e):
        if page.theme_mode == ft.ThemeMode.DARK:
            page.theme_mode = ft.ThemeMode.LIGHT
            page.bgcolor = ft.Colors.BLUE_GREY_50 
            theme_btn.content.value = "🌙" 
        else:
            page.theme_mode = ft.ThemeMode.DARK
            page.bgcolor = "#06090e"
            theme_btn.content.value = "☀️" 
        page.update()

    btn_stop = ft.ElevatedButton(
        content=ft.Text(texts[current_lang]["btn_stop"], weight=ft.FontWeight.BOLD),
        icon="stop",
        on_click=stop_search_click,
        style=ft.ButtonStyle(color=ft.Colors.GREY_700, bgcolor=ft.Colors.TRANSPARENT, elevation=0),
        disabled=True 
    )

    theme_btn = ft.TextButton(content=ft.Text("☀️", size=20), on_click=toggle_theme, tooltip="تغيير المظهر / Toggle Theme")
    lang_btn = ft.TextButton(content=ft.Text(texts[current_lang]["lang_btn_text"], weight=ft.FontWeight.BOLD), icon="language", on_click=toggle_language, style=ft.ButtonStyle(color=ft.Colors.BLUE_500))
    
    right_controls = ft.Row([theme_btn, lang_btn])
    top_row = ft.Row([btn_stop, right_controls], alignment=ft.MainAxisAlignment.SPACE_BETWEEN)

    def process_files(onedrive_path, references):
        nonlocal stop_search_flag
        wb = None
        try:
            status_text.value = "جاري قراءة مسارات الملفات... (لحظات من فضلك)"
            page.update()

            all_files = {}
            for root_dir, dirs, files in os.walk(onedrive_path):
                if stop_search_flag:
                    break
                for file in files:
                    file_lower = file.lower()
                    if file_lower.endswith(('.pdf', '.dwg', '.dxf')):
                        all_files[file_lower] = os.path.join(root_dir, file)

            if stop_search_flag:
                return

            total_refs = len(references)
            results = []
            
            for i, ref in enumerate(references):
                if stop_search_flag:
                    break
                    
                status_text.value = f"تم فحص {i + 1} من {total_refs} ريفرنس..."
                page.update()
                
                ref_str = str(ref).lower()
                
                alt1 = ref_str.split('-')[0] 
                alt2 = re.sub(r'x[a-z]$', 'x0', alt1) 
                
                found_pdf = False
                found_dxf = False
                pdf_link = ""
                dxf_link = ""
                pdf_is_alt = False
                dxf_is_alt = False

                for fname, fpath in all_files.items():
                    if exact_match(ref_str, fname):
                        if fname.endswith(('.pdf', '.dwg')) and not found_pdf:
                            found_pdf = True
                            pdf_link = fpath
                        elif fname.endswith('.dxf') and not found_dxf:
                            found_dxf = True
                            dxf_link = fpath
                    if found_pdf and found_dxf:
                        break

                if (not found_pdf or not found_dxf) and alt1 != ref_str:
                    for fname, fpath in all_files.items():
                        if exact_match(alt1, fname):
                            if not found_pdf and fname.endswith(('.pdf', '.dwg')):
                                found_pdf = True
                                pdf_link = fpath
                                pdf_is_alt = True
                            elif not found_dxf and fname.endswith('.dxf'):
                                found_dxf = True
                                dxf_link = fpath
                                dxf_is_alt = True
                        if found_pdf and found_dxf:
                            break 
                            
                if (not found_pdf or not found_dxf) and alt2 != alt1:
                    for fname, fpath in all_files.items():
                        if exact_match(alt2, fname):
                            if not found_pdf and fname.endswith(('.pdf', '.dwg')):
                                found_pdf = True
                                pdf_link = fpath
                                pdf_is_alt = True
                            elif not found_dxf and fname.endswith('.dxf'):
                                found_dxf = True
                                dxf_link = fpath
                                dxf_is_alt = True
                        if found_pdf and found_dxf:
                            break

                pdf_status = "Done (Alt)" if pdf_is_alt else ("Done" if found_pdf else "Missing")
                pdf_cell = f'=HYPERLINK("{pdf_link}", "Open PDF/DWG")' if found_pdf and len(pdf_link) <= 250 else pdf_link
                dxf_status = "Done (Alt)" if dxf_is_alt else ("Done" if found_dxf else "Missing")
                dxf_cell = f'=HYPERLINK("{dxf_link}", "Open DXF")' if found_dxf and len(dxf_link) <= 250 else dxf_link

                results.append([ref, pdf_status, dxf_status, pdf_cell, dxf_cell, pdf_is_alt, dxf_is_alt])

            if stop_search_flag:
                return

            wb = xw.Book()
            ws = wb.sheets.active
            headers = ["Reference", "PDF/DWG Status", "DXF Status", "PDF/DWG LINK", "DXF LINK"]
            ws.range("A1").value = headers
            ws.range("A1:E1").font.bold = True
            ws.range("A1:E1").color = (0, 112, 60)
            ws.range("A1:E1").font.color = (255, 255, 255)

            if results:
                ws.range("A2").value = [r[:5] for r in results]
                
                for idx, row in enumerate(results):
                    pdf_alt = row[5]
                    dxf_alt = row[6]
                    excel_row = idx + 2
                    
                    if pdf_alt:
                        ws.range(f"B{excel_row}").color = (255, 255, 0)
                        ws.range(f"B{excel_row}").font.color = (0, 0, 0)
                        ws.range(f"D{excel_row}").color = (255, 255, 0)
                        ws.range(f"D{excel_row}").font.color = (0, 0, 0)
                        
                    if dxf_alt:
                        ws.range(f"C{excel_row}").color = (255, 255, 0)
                        ws.range(f"C{excel_row}").font.color = (0, 0, 0)
                        ws.range(f"E{excel_row}").color = (255, 255, 0)
                        ws.range(f"E{excel_row}").font.color = (0, 0, 0)

            ws.autofit()
            wb.app.activate(steal_focus=True)

        except Exception as e:
            if not stop_search_flag: 
                show_snack(texts[current_lang]["snack_error"].format(str(e)), is_error=True)
            
        finally:
            btn_check.disabled = False
            btn_stop.style = ft.ButtonStyle(color=ft.Colors.GREY_700, bgcolor=ft.Colors.TRANSPARENT, elevation=0)
            btn_stop.disabled = True 
            pr.visible = False
            status_text.visible = False
            
            if stop_search_flag:
                show_snack(texts[current_lang]["snack_stopped"], is_error=True)
            else:
                try:
                    sound_file = get_resource_path("assets/success.wav")
                    winsound.PlaySound(sound_file, winsound.SND_FILENAME | winsound.SND_ASYNC)
                except Exception:
                    pass 
                show_snack(texts[current_lang]["snack_success"].format(len(references)))
            
            page.update()

    def start_checking(e):
        nonlocal stop_search_flag
        stop_search_flag = False 
        
        refs_text = txt_refs.value
        if not refs_text or not refs_text.strip():
            show_snack(texts[current_lang]["snack_empty"], is_error=True)
            return

        references = [r.strip() for r in refs_text.split('\n') if r.strip()]
        onedrive_path = os.environ.get('ONEDRIVE')
        
        if not onedrive_path:
            show_snack(texts[current_lang]["snack_no_drive"], is_error=True)
            return

        btn_check.disabled = True
        btn_stop.style = ft.ButtonStyle(color=ft.Colors.WHITE, bgcolor=ft.Colors.RED_600, elevation=5)
        btn_stop.disabled = False
        
        pr.visible = True
        status_text.value = texts[current_lang]["status_searching"].format(len(references))
        status_text.visible = True
        page.update()
        
        page.run_thread(process_files, onedrive_path, references)

    btn_check = ft.ElevatedButton(
        content=ft.Text(texts[current_lang]["btn_search"], size=16, weight=ft.FontWeight.BOLD), 
        icon="search", 
        on_click=start_checking,
        style=ft.ButtonStyle(color=ft.Colors.WHITE, bgcolor=ft.Colors.BLUE_700, shape=ft.RoundedRectangleBorder(radius=10), padding=20),
        width=400,
        height=55
    )

    credits_section = ft.Row([
        ft.Container(content=ft.Column([
            ft.Row([ft.Text("Powered By :", size=13, color=ft.Colors.GREEN_600, weight=ft.FontWeight.BOLD), ft.Text("Ihab AbdEl-Maksoud Amin", size=13, color=ft.Colors.BLUE_500, weight=ft.FontWeight.BOLD)], spacing=5),
            ft.Row([ft.Text("Designed By :", size=13, color=ft.Colors.GREEN_600, weight=ft.FontWeight.BOLD), ft.Text("Zyad Ihab AbdEl-Maksoud", size=13, color=ft.Colors.BLUE_500, weight=ft.FontWeight.BOLD)], spacing=5),
        ], spacing=2))
    ], alignment=ft.MainAxisAlignment.START)

    page.add(
        top_row, logo_image, ft.Divider(height=10, color=ft.Colors.TRANSPARENT), txt_refs,
        ft.Divider(height=10, color=ft.Colors.TRANSPARENT), status_text, pr,
        ft.Divider(height=10, color=ft.Colors.TRANSPARENT), btn_check,
        ft.Divider(height=15, color=ft.Colors.TRANSPARENT), credits_section 
    )

assets_path = get_resource_path("assets")
ft.app(target=main, assets_dir=assets_path)
import tkinter as tk
from tkinter import messagebox, filedialog
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
import os
import shutil
from datetime import datetime

class AppendixDialog:
    """Диалог добавления приложения (без количества листов)"""
    def __init__(self, parent, title="", text=""):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Приложение")
        self.dialog.geometry("550x450")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        self.result = None

        tk.Label(self.dialog, text="Заголовок приложения (будет по центру):").pack(pady=5)
        self.title_entry = tk.Entry(self.dialog, width=60)
        self.title_entry.insert(0, title)
        self.title_entry.pack()

        tk.Label(self.dialog, text="Текст приложения:").pack(pady=5)
        self.text_text = tk.Text(self.dialog, height=15, width=60)
        self.text_text.insert("1.0", text)
        self.text_text.pack()

        btn_frame = tk.Frame(self.dialog)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="OK", command=self.ok).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Отмена", command=self.cancel).pack(side=tk.LEFT, padx=5)

    def ok(self):
        title = self.title_entry.get().strip()
        text = self.text_text.get("1.0", tk.END).strip()
        if not title:
            messagebox.showwarning("Ошибка", "Заголовок не может быть пустым")
            return
        self.result = {"title": title, "text": text}
        self.dialog.destroy()

    def cancel(self):
        self.dialog.destroy()


class LetterWithAppendices:
    def __init__(self, root):
        self.root = root
        self.root.title("Генератор писем с приложениями")
        self.root.geometry("850x870")          
        self.root.resizable(False, False)

        self.template_path = ""
        self.appendices = []

        self.create_widgets()

    def create_widgets(self):
        canvas = tk.Canvas(self.root)
        scrollbar = tk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        self.entries = {}
        row = 0

        tk.Label(scrollable_frame, text="ОТПРАВИТЕЛЬ", font=("Arial", 10, "bold")).grid(row=row, column=0, columnspan=2, pady=5, sticky="w")
        row += 1
        sender_fields = [
            ("Название организации:", "org_name", "АО «Пролетарский прогресс»"),
            ("Адрес:", "org_address", "г. Москва, ул. Ленина, д. 5"),
            ("Телефон:", "org_phone", "(495) 123-45-67"),
            ("E-mail:", "org_email", "info@example.ru"),
        ]
        for label_text, key, default in sender_fields:
            tk.Label(scrollable_frame, text=label_text).grid(row=row, column=0, sticky="w", pady=2, padx=5)
            entry = tk.Entry(scrollable_frame, width=60)
            entry.insert(0, default)
            entry.grid(row=row, column=1, pady=2, padx=5)
            self.entries[key] = entry
            row += 1

        # ---- Реквизиты письма ----
        row += 1
        tk.Label(scrollable_frame, text="РЕКВИЗИТЫ ПИСЬМА", font=("Arial", 10, "bold")).grid(row=row, column=0, columnspan=2, pady=5, sticky="w")
        row += 1
        # Дата
        tk.Label(scrollable_frame, text="Дата (дд.мм.гггг):").grid(row=row, column=0, sticky="w", pady=2, padx=5)
        self.date_entry = tk.Entry(scrollable_frame, width=60)
        self.date_entry.insert(0, datetime.now().strftime("%d.%m.%Y"))
        self.date_entry.grid(row=row, column=1, pady=2, padx=5)
        row += 1
        # Исходящий номер
        tk.Label(scrollable_frame, text="Исходящий номер (после 'от'):").grid(row=row, column=0, sticky="w", pady=2, padx=5)
        self.out_number_entry = tk.Entry(scrollable_frame, width=60)
        self.out_number_entry.insert(0, "12/05")
        self.out_number_entry.grid(row=row, column=1, pady=2, padx=5)
        row += 1
        # Входящий номер
        tk.Label(scrollable_frame, text="Входящий номер (после 'На №'):").grid(row=row, column=0, sticky="w", pady=2, padx=5)
        self.in_number_entry = tk.Entry(scrollable_frame, width=60)
        self.in_number_entry.insert(0, "45")
        self.in_number_entry.grid(row=row, column=1, pady=2, padx=5)
        row += 1

        # ---- Адресат ----
        row += 1
        tk.Label(scrollable_frame, text="АДРЕСАТ", font=("Arial", 10, "bold")).grid(row=row, column=0, columnspan=2, pady=5, sticky="w")
        row += 1
        recipient_fields = [
            ("Должность адресата (в дательном падеже):", "recipient_position", "Генеральному директору"),
            ("Организация адресата:", "recipient_company", "ООО «Недвижимость»"),
            ("ФИО адресата (в дательном падеже):", "recipient_name", "Иванову И.И."),
        ]
        for label_text, key, default in recipient_fields:
            tk.Label(scrollable_frame, text=label_text).grid(row=row, column=0, sticky="w", pady=2, padx=5)
            entry = tk.Entry(scrollable_frame, width=60)
            entry.insert(0, default)
            entry.grid(row=row, column=1, pady=2, padx=5)
            self.entries[key] = entry
            row += 1

        row += 1
        tk.Label(scrollable_frame, text="ТЕКСТ ПИСЬМА", font=("Arial", 10, "bold")).grid(row=row, column=0, columnspan=2, pady=5, sticky="w")
        row += 1
        tk.Label(scrollable_frame, text="Содержание письма (можно использовать {RECIPIENT_NAME}):").grid(row=row, column=0, sticky="nw", pady=2, padx=5)
        default_body = """Уважаемому {RECIPIENT_NAME}!

Настоящим письмом направляем вам информацию и прилагаем необходимые документы.

Просим рассмотреть и сообщить о решении."""
        self.body_text = tk.Text(scrollable_frame, height=8, width=55)   
        self.body_text.insert("1.0", default_body)
        self.body_text.grid(row=row, column=1, pady=2, padx=5)
        row += 1

        # ---- Приложения ----
        row += 1
        tk.Label(scrollable_frame, text="ПРИЛОЖЕНИЯ", font=("Arial", 10, "bold")).grid(row=row, column=0, columnspan=2, pady=5, sticky="w")
        row += 1

        self.appendices_listbox = tk.Listbox(scrollable_frame, height=5, width=70)
        self.appendices_listbox.grid(row=row, column=0, columnspan=2, pady=5, padx=5)
        row += 1

        btn_frame = tk.Frame(scrollable_frame)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=5)
        tk.Button(btn_frame, text="Добавить приложение", command=self.add_appendix).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Редактировать", command=self.edit_appendix).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Удалить", command=self.delete_appendix).pack(side=tk.LEFT, padx=5)
        row += 1

        row += 1
        tk.Label(scrollable_frame, text="ПОДПИСЬ И ИСПОЛНИТЕЛЬ", font=("Arial", 10, "bold")).grid(row=row, column=0, columnspan=2, pady=5, sticky="w")
        row += 1
        sign_fields = [
            ("Должность подписанта:", "signer_position", "Генеральный директор"),
            ("ФИО подписанта:", "signer_name", "Петров П.П."),
            ("Исполнитель (ФИО):", "executor_name", "Сидорова А.А."),
            ("Телефон исполнителя:", "executor_phone", "(495) 999-88-77"),
        ]
        for label_text, key, default in sign_fields:
            tk.Label(scrollable_frame, text=label_text).grid(row=row, column=0, sticky="w", pady=2, padx=5)
            entry = tk.Entry(scrollable_frame, width=60)
            entry.insert(0, default)
            entry.grid(row=row, column=1, pady=2, padx=5)
            self.entries[key] = entry
            row += 1

        row += 1
        self.template_btn = tk.Button(scrollable_frame, text="1. Выбрать шаблон DOCX", command=self.select_template, bg="lightgray")
        self.template_btn.grid(row=row, column=0, columnspan=2, pady=10)

        self.template_label = tk.Label(scrollable_frame, text="Шаблон не выбран", fg="red")
        self.template_label.grid(row=row+1, column=0, columnspan=2)

        self.generate_btn = tk.Button(scrollable_frame, text="2. Создать письмо", command=self.generate_letter, bg="lightblue", state=tk.DISABLED)
        self.generate_btn.grid(row=row+2, column=0, columnspan=2, pady=10)

        self.status_label = tk.Label(scrollable_frame, text="Готов", fg="green")
        self.status_label.grid(row=row+3, column=0, columnspan=2)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.update_appendix_listbox()

    def update_appendix_listbox(self):
        self.appendices_listbox.delete(0, tk.END)
        for idx, app in enumerate(self.appendices, 1):
            self.appendices_listbox.insert(tk.END, f"{idx}. {app['title']}")

    def add_appendix(self):
        dialog = AppendixDialog(self.root)
        self.root.wait_window(dialog.dialog)
        if dialog.result:
            self.appendices.append(dialog.result)
            self.update_appendix_listbox()

    def edit_appendix(self):
        selection = self.appendices_listbox.curselection()
        if not selection:
            messagebox.showinfo("Редактирование", "Выберите приложение для редактирования")
            return
        idx = selection[0]
        app = self.appendices[idx]
        dialog = AppendixDialog(self.root, title=app["title"], text=app["text"])
        self.root.wait_window(dialog.dialog)
        if dialog.result:
            self.appendices[idx] = dialog.result
            self.update_appendix_listbox()

    def delete_appendix(self):
        selection = self.appendices_listbox.curselection()
        if not selection:
            return
        idx = selection[0]
        del self.appendices[idx]
        self.update_appendix_listbox()

    def select_template(self):
        file_path = filedialog.askopenfilename(filetypes=[("DOCX файлы", "*.docx")])
        if file_path:
            self.template_path = file_path
            self.template_label.config(text=f"Шаблон: {os.path.basename(file_path)}", fg="green")
            self.generate_btn.config(state=tk.NORMAL)

    def replace_placeholders(self, doc, replacements):
        for paragraph in doc.paragraphs:
            self.replace_in_paragraph(paragraph, replacements)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self.replace_in_paragraph(paragraph, replacements)

    def replace_in_paragraph(self, paragraph, replacements):
        full_text = paragraph.text
        new_text = full_text
        for placeholder, value in replacements.items():
            if value is None:
                value = ""
            new_text = new_text.replace(placeholder, value)
        if new_text != full_text:
            paragraph.clear()
            paragraph.add_run(new_text)

    def generate_appendix_list_text(self):
        if not self.appendices:
            return ""
        lines = ["Приложение:"]
        for idx, app in enumerate(self.appendices, 1):
            lines.append(f"{idx}. {app['title']}")
        return "\n".join(lines)

    def insert_appendices_into_doc(self, doc):
        for paragraph in doc.paragraphs:
            if "{APPENDICES_CONTENT}" in paragraph.text:
                p = paragraph._element
                p.getparent().remove(p)
                break

        if not self.appendices:
            return

        doc.add_page_break()

        for idx, app in enumerate(self.appendices, 1):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if len(self.appendices) == 1:
                p.add_run("Приложение")
            else:
                p.add_run(f"Приложение {idx}")

            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(app["title"])
            run.bold = True
            run.font.size = Pt(14)

            lines = app["text"].splitlines()
            for line in lines:
                if line.strip():
                    p = doc.add_paragraph(line.strip())
                else:
                    doc.add_paragraph()

            if idx < len(self.appendices):
                doc.add_page_break()

    def generate_letter(self):
        if not self.template_path:
            messagebox.showerror("Ошибка", "Сначала выберите шаблон!")
            return

        replacements = {
            "{ORG_NAME}": self.entries["org_name"].get().strip(),
            "{ORG_ADDRESS}": self.entries["org_address"].get().strip(),
            "{ORG_PHONE}": self.entries["org_phone"].get().strip(),
            "{ORG_EMAIL}": self.entries["org_email"].get().strip(),
            "{DATE}": self.date_entry.get().strip(),
            "{OUT_NUMBER}": self.out_number_entry.get().strip(),
            "{IN_NUMBER}": self.in_number_entry.get().strip(),
            "{RECIPIENT_POSITION}": self.entries["recipient_position"].get().strip(),
            "{RECIPIENT_COMPANY}": self.entries["recipient_company"].get().strip(),
            "{RECIPIENT_NAME}": self.entries["recipient_name"].get().strip(),
            "{SIGNER_POSITION}": self.entries["signer_position"].get().strip(),
            "{SIGNER_NAME}": self.entries["signer_name"].get().strip(),
            "{EXECUTOR_NAME}": self.entries["executor_name"].get().strip(),
            "{EXECUTOR_PHONE}": self.entries["executor_phone"].get().strip(),
        }

        body_text = self.body_text.get("1.0", tk.END).strip()
        body_text = body_text.replace("{RECIPIENT_NAME}", replacements["{RECIPIENT_NAME}"])
        replacements["{BODY}"] = body_text

        if not replacements["{BODY}"]:
            messagebox.showwarning("Предупреждение", "Текст письма не может быть пустым")
            return

        replacements["{APPENDICES_LIST}"] = self.generate_appendix_list_text()

        output_dir = os.path.join(os.path.expanduser("~"), "Desktop")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(output_dir, f"Letter_{timestamp}.docx")

        try:
            shutil.copy2(self.template_path, output_path)
            doc = Document(output_path)

            self.replace_placeholders(doc, replacements)
            self.insert_appendices_into_doc(doc)

            doc.save(output_path)

            self.status_label.config(text=f"Готово: {output_path}", fg="blue")
            messagebox.showinfo("Успех", f"Письмо создано!\n{output_path}")
        except Exception as e:
            self.status_label.config(text="Ошибка!", fg="red")
            messagebox.showerror("Ошибка", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = LetterWithAppendices(root)
    root.mainloop()
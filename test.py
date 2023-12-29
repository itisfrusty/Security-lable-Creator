# import tkinter as tk
# from datetime import datetime
# import openpyxl
# import rembg
# import os
# from tkinter import filedialog
# from PIL import Image, ImageDraw, ImageFont, ImageTk
#
# def overlay_text_and_image(image_path, text_data, overlay_image_path, position_data, output_path, photo_path):
#     base_image = Image.open(image_path)
#     overlay_image = Image.open(overlay_image_path)
#
#     for text, position in zip(text_data, position_data):
#         draw = ImageDraw.Draw(base_image)
#         font = ImageFont.truetype("Play-Bold.ttf", 40)  #Customize the font and size here
#         draw.text(position, text, fill="black", font=font)
#
#     new_size = (250, 330)
#     resized_overlay_image = overlay_image.resize(new_size)
#     base_image.paste(resized_overlay_image, (20, 150))
#     base_image.save(output_path)
#
# def remove_background(image_path, output_path):
#     input_image = Image.open(image_path)
#     output_image = rembg.remove(input_image)
#
#     # Convert the output image to RGBA mode
#     output_image = output_image.convert("RGBA")
#
#     # Create a new image with a white background
#     width, height = output_image.size
#     white_background = Image.new("RGBA", (width, height), (255, 255, 255, 255))
#
#     # Composite the output image on the white background
#     result = Image.alpha_composite(white_background, output_image)
#
#     result.save(output_path, "PNG")
#
# def add_photo():
#     image_to_overlay_path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg;*.jpeg;*.png")])
#     if image_to_overlay_path:
#
#         # Удаляем фон и сохраняем результат
#         processed_image_path = "processed_image.png"
#         remove_background(image_to_overlay_path,processed_image_path)
#
#         photo_path.set(processed_image_path)
#         photo = Image.open(processed_image_path)
#         photo.thumbnail((200, 200))  # Resize the photo to fit the GUI
#         photo = ImageTk.PhotoImage(photo)
#         photo_label.config(image=photo)
#         photo_label.image = photo  # Keep a reference to avoid garbage collection
#
#         #Обновить размер окна
#         new_width = 550
#         new_height = 670
#         app.geometry(f"{new_width}x{new_height}")
#
# # def create_pass_on_a4(pass_image_path):
# #     # Размер листа A4 в пикселях (300 DPI)
# #     A4_WIDTH = 2480
# #     A4_HEIGHT = 3508
# #
# #     # Размер изображения на пропуске
# #     PASS_WIDTH = 827
# #     PASS_HEIGHT = 531
# #
# #     # Размеры отступов от краев листа A4 (1 дюйм = 25.4 мм)
# #     LEFT_MARGIN = 150  # Отступ слева
# #     TOP_MARGIN = 150  # Отступ сверху
# #
# #     # Создаем новое изображение с размерами листа A4
# #     output_image = Image.new("RGB", (A4_WIDTH, A4_HEIGHT), "white")
# #
# #     # Загружаем изображение пропуска
# #     pass_image = Image.open(pass_image_path)
# #     pass_image = pass_image.resize((PASS_WIDTH, PASS_HEIGHT))
# #
# #     # Вставляем изображение пропуска на лист A4 с учетом отступов
# #     output_image.paste(pass_image, (int(LEFT_MARGIN), int(TOP_MARGIN)))
# #
# #     # Получаем путь к рабочему столу
# #     desktop_path = os.path.expanduser("~/Desktop")
# #
# #     # Сохраняем получившееся изображение на рабочем столе
# #     output_image_path = os.path.join(desktop_path, "propusk_a4.jpg")
# #     output_image.save(output_image_path)
#
# def clear_photo_label():
#     blank_image = Image.new("RGB", (0, 0), "white")  # Создаем пустое изображение
#     blank_photo = ImageTk.PhotoImage(blank_image)  # Создаем PhotoImage из пустого изображения
#     photo_label.config(image=blank_photo)  # Присваиваем виджету пустое изображение
#     photo_label.image = blank_photo
#     app.geometry("550x470")
#
# def add_to_excel(text_data):
#     excel_filename = "data.xlsx"
#
#     # Create or open the workbook
#     if os.path.exists(excel_filename):
#         workbook = openpyxl.load_workbook(excel_filename)
#         sheet = workbook.active
#     else:
#         workbook = openpyxl.Workbook()
#         sheet = workbook.active
#         sheet.append(["Пропуск №", "Фамилия", "Имя", "Отчество", "Табельный №"])
#
#     empty_row = None
#     for row_num in range(2, sheet.max_row + 1):
#         if all(cell.value is None or cell.value == "" for cell in sheet[row_num]):
#             empty_row = row_num
#             break
#
#     if not empty_row:
#         empty_row = sheet.max_row + 1
#         sheet.append([None] * 5)
#
#     # Write text_data to the empty row
#     for col_num, value in enumerate(text_data, start=1):
#         sheet.cell(row=empty_row, column=col_num, value=value)
#
#     # Save the workbook
#     workbook.save(excel_filename)
#
# def save_data():
#     text_data = [entry.get() for entry in text_entries]
#
#     # Check if at least one field is filled
#     if any(data.strip() for data in text_data):
#         add_to_excel(text_data)
#         show_save_success_message()
#     else:
#         message = "Пожалуйста, заполните хотя бы\n одно поле для сохранения данных :("
#         popup = tk.Toplevel()
#         popup.geometry("350x60")
#         popup.title("Ошибка")
#         label = tk.Label(popup, text=message, font=("Play-Bold.ttf", 14))
#         label.pack()
#
# def show_message():
#     #Create popup message
#     message = "Пропуск создан!\nДанные записаны в документ! "
#     popup = tk.Toplevel()
#     popup.geometry("300x60")
#     popup.title("Сообщение")
#     label = tk.Label(popup, text=message, font=("Play-Bold.ttf", 14))
#     label.pack()
#
# def show_save_success_message():
#     message = "Данные успешно сохранены :)"
#     popup = tk.Toplevel()
#     popup.geometry("300x40")
#     popup.title("Сообщение")
#     label = tk.Label(popup, text=message, font=("Play-Bold.ttf", 14))
#     label.pack()
#
# def submit():
#     text_data = [entry.get() for entry in text_entries]
#     position_data = [(685, 162), (500, 260), (500, 303), (500, 346), (565, 433)]  # Customize the positions here
#     current_time = datetime.now().strftime('%Y%m%d%H%M%S')
#
#     # Create the folder "propuska" at the root of the application
#     propuska_folder = os.path.join(os.path.dirname(__file__), "propuska")
#     os.makedirs(propuska_folder, exist_ok=True)
#
#     output_filename = f"propusk_{current_time}.jpg"
#     output_filepath = os.path.join(propuska_folder, output_filename)
#
#     overlay_text_and_image("pass.jpg", text_data, photo_path.get(), position_data, output_filepath, photo_path.get())
#
#     # Clear the text entries
#     for entry in text_entries:
#         entry.delete(0, tk.END)
#
#     # Clear the photo label
#     clear_photo_label()
#     show_message()
#
# # Функция для изменения размера шрифта
# def change_font_size(event):
#     new_font = ("Play-Bold.ttf", 14)  # Здесь задайте новый размер шрифта
#     for entry in text_entries:
#         entry.config(font=new_font)
#
# app = tk.Tk()
# app.title("Создание нового пропуска")
# app.geometry("550x470")  # Задаем размер главного окна
#
# # Создание полей для ввода текста
# label_names = ["Номер пропуска:", "Фамилия:", "Имя:", "Отчество:", "Табельный номер:"]
# text_entries = []
# for i, label_name in enumerate(label_names):
#     label = tk.Label(app, text=label_name, font=("Play-Bold.ttf", 14))
#     label.grid(row=i, column=0, padx=5, pady=5, sticky=tk.W)
#     entry = tk.Entry(app, width=30)
#     entry.grid(row=i, column=1, padx=5, pady=5)
#     text_entries.append(entry)
#
# # Кнопка для добавления фотографии
# photo_button = tk.Button(app, text="Добавить фотографию", command=add_photo, width=20, height=1, font=("Play-Bold.ttf", 14))
# photo_button.grid(row=len(label_names), column=0, columnspan=2, padx=5, pady=10)
#
# # Поле для отображения фотографии
# photo_path = tk.StringVar()
# photo_label = tk.Label(app)
# photo_label.grid(row=len(label_names) + 1, column=0, columnspan=2, padx=5, pady=10)
#
# # Кнопка "Сохранить данные"
# save_button = tk.Button(app, text="Сохранить данные", command=save_data, width=20, height=1, font=("Play-Bold.ttf", 14))
# save_button.grid(row=len(label_names) + 2, column=0, columnspan=2, padx=5, pady=10)
#
# # Кнопка "Создать пропуск"
# submit_button = tk.Button(app, text="Создать пропуск", command=submit, width=20, height=2, font=("Play-Bold.ttf", 14))
# submit_button.grid(row=len(label_names) + 3, column=0, columnspan=2, padx=5, pady=10)
#
# # Привязка функции к событию изменения размера окна (можно настроить по своему)
# app.bind("<Configure>", change_font_size)  # Функция для изменения размера шрифта
#
# app.mainloop()

    for i in range(len(buttons)):
            # Tính toán vị trí dòng và cột dựa trên số cột tối đa (max_columns)
            row = i // max_columns
            col = i % max_columns

            button_text = buttons[i]
            btn = tk.Button(window, text=button_text, padx=20, pady=20, command=lambda text=button_text: button_click(text))
            btn.grid(row=row, column=col)

    # Số cột tối đa bạn muốn hiển thị
        max_columns = 4
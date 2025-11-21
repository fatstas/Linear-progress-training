import tkinter as tk
from tkinter import ttk, messagebox


class RMCalculator:
    def __init__(self, parent):
        self.parent = parent

    def create_calculator_tab(self):
        """Создает вкладку с калькулятором 1ПМ"""

        calculator_frame = ttk.Frame(self.parent)

        # Основной фрейм калькулятора
        main_calc_frame = ttk.Frame(calculator_frame, padding="15")
        main_calc_frame.pack(fill="both", expand=True)

        # Ввод данных
        input_frame = ttk.LabelFrame(main_calc_frame, text="Входные данные", padding="10")
        input_frame.pack(fill="x", pady=(0, 15))

        # Вес
        ttk.Label(input_frame, text="Вес (кг):", font=("Arial", 11)).grid(row=0, column=0, sticky=tk.W, pady=8)
        self.rm_weight_entry = ttk.Entry(input_frame, width=12, font=("Arial", 11))
        self.rm_weight_entry.grid(row=0, column=1, sticky=tk.W, pady=8, padx=(10, 30))
        self.rm_weight_entry.insert(0, "100")

        # Количество повторений
        ttk.Label(input_frame, text="Повторения:", font=("Arial", 11)).grid(row=0, column=2, sticky=tk.W, pady=8)
        self.rm_reps_entry = ttk.Entry(input_frame, width=12, font=("Arial", 11))
        self.rm_reps_entry.grid(row=0, column=3, sticky=tk.W, pady=8, padx=(10, 30))
        self.rm_reps_entry.insert(0, "5")

        # Кнопка расчета
        ttk.Button(input_frame, text="🎯 Рассчитать 1ПМ",
                   command=self.calculate_1rm).grid(row=0, column=4, sticky=tk.W, pady=8, padx=(20, 0))

        # Результаты
        results_frame = ttk.LabelFrame(main_calc_frame, text="Результаты расчетов", padding="10")
        results_frame.pack(fill="both", expand=True)

        # Таблица результатов
        columns = ("Формула", "1ПМ (кг)", "Разница")
        self.rm_tree = ttk.Treeview(results_frame, columns=columns, show="headings", height=12)

        # Заголовки
        self.rm_tree.heading("Формула", text="Формула")
        self.rm_tree.heading("1ПМ (кг)", text="1ПМ (кг)")
        self.rm_tree.heading("Разница", text="Разница")

        # Ширина колонок
        self.rm_tree.column("Формула", width=150)
        self.rm_tree.column("1ПМ (кг)", width=100)
        self.rm_tree.column("Разница", width=100)

        self.rm_tree.pack(fill="both", expand=True)

        # Прокрутка для таблицы
        scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=self.rm_tree.yview)
        self.rm_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        return calculator_frame

    def calculate_1rm(self):
        """Рассчитывает 1ПМ по разным формулам"""
        try:
            weight = float(self.rm_weight_entry.get())
            reps = int(self.rm_reps_entry.get())

            if weight <= 0 or reps <= 0:
                messagebox.showerror("Ошибка", "Вес и повторения должны быть положительными числами")
                return

            # Очищаем предыдущие результаты
            for item in self.rm_tree.get_children():
                self.rm_tree.delete(item)

            # Рассчитываем по всем формулам
            formulas = {
                "Эпли": self.epley_1rm,
                "Бжицки": self.brzycki_1rm,
                "Лэндер": self.lander_1rm,
                "Ломбарди": self.lombardi_1rm,
                "Мэйхью": self.mayhew_1rm,
                "О'Коннор": self.oconnor_1rm,
                "Ватан": self.wathan_1rm
            }

            results = []
            for name, formula in formulas.items():
                rm = formula(weight, reps)
                results.append((name, rm))

            # Сортируем по величине 1ПМ
            results.sort(key=lambda x: x[1])

            # Добавляем в таблицу
            avg_rm = sum(rm for _, rm in results) / len(results)

            for name, rm in results:
                diff = rm - avg_rm
                diff_text = f"{diff:+.1f}" if abs(diff) >= 0.1 else "0.0"
                self.rm_tree.insert("", "end", values=(name, f"{rm:.1f}", diff_text))

            # Среднее значение
            self.rm_tree.insert("", "end", values=("СРЕДНЕЕ", f"{avg_rm:.1f}", "0.0"), tags=("average",))
            self.rm_tree.tag_configure("average", background="lightgray", font=("Arial", 10, "bold"))

        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректные числовые значения")

    # Формулы расчета 1ПМ
    def epley_1rm(self, weight, reps):
        """Формула Эпли"""
        return weight * (1 + reps / 30)

    def brzycki_1rm(self, weight, reps):
        """Формула Бжицки"""
        return weight * (36 / (37 - reps))

    def lander_1rm(self, weight, reps):
        """Формула Лэндера"""
        return (100 * weight) / (101.3 - 2.67123 * reps)

    def lombardi_1rm(self, weight, reps):
        """Формула Ломбарди"""
        return weight * (reps ** 0.10)

    def mayhew_1rm(self, weight, reps):
        """Формула Мэйхью"""
        return (100 * weight) / (52.2 + 41.9 * (2.71828 ** (-0.055 * reps)))

    def oconnor_1rm(self, weight, reps):
        """Формула О'Коннора"""
        return weight * (1 + reps * 0.025)

    def wathan_1rm(self, weight, reps):
        """Формула Ватана"""
        return (100 * weight) / (48.8 + 53.8 * (2.71828 ** (-0.075 * reps)))
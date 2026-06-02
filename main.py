import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os

APP_NAME = "Label Generator"
GROUP_COLUMN = "رمز بلد عنوان المستلم (country code)"
DUPLICATE_COLUMN = "عدد البوالص"


class LabelGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("1000x720")

        self.columns = []
        self.entries = {}
        self.orders = []

        # ---------- Header ----------
        ttk.Label(
            root,
            text=APP_NAME,
            font=("Segoe UI", 16, "bold")
        ).pack(pady=10)

        ttk.Button(
            root,
            text="Load Excel Template",
            command=self.load_excel
        ).pack(pady=5)

        # ---------- Form ----------
        self.form_frame = tk.LabelFrame(root, text="Order Entry")
        self.form_frame.pack(fill="x", padx=10, pady=5)

        # ---------- Buttons ----------
        btns = tk.Frame(root)
        btns.pack(pady=5)

        ttk.Button(btns, text="Add Order", command=self.add_order).pack(side="left", padx=5)
        ttk.Button(btns, text="Delete Selected Order", command=self.delete_order).pack(side="left", padx=5)
        ttk.Button(btns, text="Export All Orders", command=self.export_all).pack(side="left", padx=5)

        # ---------- Orders Table ----------
        self.table_frame = tk.LabelFrame(root, text="Orders Table")
        self.table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.tree = ttk.Treeview(self.table_frame, show="headings")
        self.tree.pack(fill="both", expand=True, side="left")

        scrollbar = ttk.Scrollbar(
            self.table_frame,
            orient="vertical",
            command=self.tree.yview
        )
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.status = ttk.Label(root, text="Orders: 0")
        self.status.pack(pady=5)

    # ---------- Load Excel ----------
    def load_excel(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not path:
            return

        try:
            df = pd.read_excel(path)
            self.columns = list(df.columns)

            # Clear form
            for widget in self.form_frame.winfo_children():
                widget.destroy()
            self.entries.clear()

            # Build dynamic form
            for col in self.columns:
                row = tk.Frame(self.form_frame)
                row.pack(fill="x", pady=2)

                ttk.Label(
                    row,
                    text=col,
                    width=35,
                    anchor="w"
                ).pack(side="left")

                entry = ttk.Entry(row)
                entry.pack(side="right", fill="x", expand=True)
                self.entries[col] = entry

            # Setup table
            self.tree.delete(*self.tree.get_children())
            self.tree["columns"] = self.columns

            for col in self.columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=150, anchor="center")

            messagebox.showinfo(
                "Ready",
                "Template loaded.\nFill fields and click Add Order."
            )

        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ---------- Add Order ----------
    def add_order(self):
        if not self.entries:
            messagebox.showwarning("Warning", "Load Excel template first")
            return

        order = {}

        for col, entry in self.entries.items():
            value = entry.get().strip()
            if value == "":
                messagebox.showerror("Error", f"Field '{col}' is empty")
                return
            order[col] = value

        # Validate duplication column
        try:
            order[DUPLICATE_COLUMN] = int(order[DUPLICATE_COLUMN])
            if order[DUPLICATE_COLUMN] <= 0:
                raise ValueError
        except Exception:
            messagebox.showerror(
                "Error",
                f"{DUPLICATE_COLUMN} must be a positive number"
            )
            return

        self.orders.append(order)
        self.tree.insert(
            "",
            "end",
            values=[order[c] for c in self.columns]
        )

        self.status.config(text=f"Orders: {len(self.orders)}")

        # Clear form
        for entry in self.entries.values():
            entry.delete(0, tk.END)

    # ---------- Delete Order ----------
    def delete_order(self):
        selected = self.tree.selection()
        if not selected:
            return

        index = self.tree.index(selected[0])
        self.tree.delete(selected[0])
        del self.orders[index]

        self.status.config(text=f"Orders: {len(self.orders)}")

    # ---------- Export ----------
    def export_all(self):
        if not self.orders:
            messagebox.showwarning("Warning", "No orders to export")
            return

        folder = filedialog.askdirectory()
        if not folder:
            return

        try:
            all_rows = []

            for order in self.orders:
                count = int(order[DUPLICATE_COLUMN])

                for _ in range(count):
                    row = order.copy()
                    row[DUPLICATE_COLUMN] = 1   # 🔴 force to 1 on export
                    row["قيمة الشحنة"] = 1
                    all_rows.append(row)

            df = pd.DataFrame(all_rows)

            if GROUP_COLUMN not in df.columns:
                messagebox.showerror(
                    "Error",
                    f"Missing column: {GROUP_COLUMN}"
                )
                return

            grouped = df.groupby(GROUP_COLUMN)

            for country, group_df in grouped:
                safe_name = str(country).strip().replace("/", "_")
                file_path = os.path.join(folder, f"{safe_name}.xlsx")
                group_df.to_excel(file_path, index=False)

            messagebox.showinfo(
                "Success",
                "All orders exported successfully"
            )

        except Exception as e:
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = LabelGeneratorApp(root)
    root.mainloop()
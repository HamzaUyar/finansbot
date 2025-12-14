"""
Basit Tk tabanlı masaüstü arayüz.
Kullanıcıdan data.xlsx ve konsolidasyon dosyasını alır, güncellenmiş Excel'i üretir.
Tek dosya .exe olarak PyInstaller ile paketlenebilir.
"""

import threading
import sys
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

from update_konsolidasyon import run_update


def get_base_dir() -> Path:
    """Çalışan dosyanın bulunduğu klasörü döndürür (PyInstaller uyumlu)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


class KonsolidasyonApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Konsolidasyon Güncelleyici")
        self.root.geometry("720x320")
        self.root.resizable(False, False)

        base_dir = get_base_dir()

        self.data_path = tk.StringVar(value=str((base_dir / "data.xlsx").resolve()))
        self.kons_path = tk.StringVar(value=str((base_dir / "Konsolidasyon_2025_NV (1).xlsx").resolve()))
        self.output_path = tk.StringVar(
            value=str((base_dir / "Konsolidasyon_2025_NV (1)_guncel.xlsx").resolve())
        )
        self.status = tk.StringVar(value="Dosyaları seçip 'Güncellemeyi Başlat' butonuna tıklayın.")

        self._build_ui()
        self.worker: threading.Thread | None = None

    def _build_ui(self) -> None:
        padding = {"padx": 12, "pady": 8}

        tk.Label(self.root, text="1) data.xlsx (Kaynak veriler)").grid(row=0, column=0, sticky="w", **padding)
        tk.Entry(self.root, textvariable=self.data_path, width=70).grid(row=0, column=1, **padding)
        tk.Button(self.root, text="Seç...", command=self._pick_data).grid(row=0, column=2, **padding)

        tk.Label(self.root, text="2) Konsolidasyon dosyası (Şablon)").grid(row=1, column=0, sticky="w", **padding)
        tk.Entry(self.root, textvariable=self.kons_path, width=70).grid(row=1, column=1, **padding)
        tk.Button(self.root, text="Seç...", command=self._pick_kons).grid(row=1, column=2, **padding)

        tk.Label(self.root, text="3) Çıktı dosyası").grid(row=2, column=0, sticky="w", **padding)
        tk.Entry(self.root, textvariable=self.output_path, width=70).grid(row=2, column=1, **padding)
        tk.Button(self.root, text="Kaydet...", command=self._pick_output).grid(row=2, column=2, **padding)

        self.status_label = tk.Label(self.root, textvariable=self.status, fg="#444", wraplength=560, justify="left")
        self.status_label.grid(row=3, column=0, columnspan=3, sticky="w", padx=12, pady=(4, 0))

        self.run_button = tk.Button(
            self.root,
            text="Güncellemeyi Başlat",
            command=self._start_update,
            bg="#2563eb",
            fg="white",
            padx=20,
            pady=8,
            activebackground="#1d4ed8",
        )
        self.run_button.grid(row=4, column=0, columnspan=3, pady=20)

    def _pick_data(self) -> None:
        path = filedialog.askopenfilename(
            title="data.xlsx seçin",
            filetypes=[("Excel Dosyası", "*.xlsx"), ("Tüm Dosyalar", "*.*")],
        )
        if path:
            self.data_path.set(path)

    def _pick_kons(self) -> None:
        path = filedialog.askopenfilename(
            title="Konsolidasyon dosyasını seçin",
            filetypes=[("Excel Dosyası", "*.xlsx"), ("Tüm Dosyalar", "*.*")],
        )
        if path:
            self.kons_path.set(path)
            kons_path = Path(path)
            suggested = kons_path.with_name(f"{kons_path.stem}_guncel{kons_path.suffix}")
            self.output_path.set(str(suggested))

    def _pick_output(self) -> None:
        initial = Path(self.output_path.get()).name
        path = filedialog.asksaveasfilename(
            title="Çıktı dosyası",
            defaultextension=".xlsx",
            initialfile=initial,
            filetypes=[("Excel Dosyası", "*.xlsx")],
        )
        if path:
            self.output_path.set(path)

    def _start_update(self) -> None:
        # Ön kontrol
        data_file = Path(self.data_path.get()).expanduser()
        kons_file = Path(self.kons_path.get()).expanduser()
        output_file = Path(self.output_path.get()).expanduser()

        if not data_file.exists():
            messagebox.showerror("Hata", f"Data dosyası bulunamadı:\n{data_file}")
            return
        if not kons_file.exists():
            messagebox.showerror("Hata", f"Konsolidasyon dosyası bulunamadı:\n{kons_file}")
            return
        if not output_file.name.lower().endswith(".xlsx"):
            messagebox.showerror("Hata", "Çıktı dosyası .xlsx uzantılı olmalı.")
            return

        output_file.parent.mkdir(parents=True, exist_ok=True)

        self._set_running_state(True)
        self.status.set("İşlem çalışıyor, lütfen bekleyin...")

        self.worker = threading.Thread(
            target=self._run_update_thread, args=(data_file, kons_file, output_file), daemon=True
        )
        self.worker.start()

    def _run_update_thread(self, data_file: Path, kons_file: Path, output_file: Path) -> None:
        try:
            target_file, last_month = run_update(data_file, kons_file, output_file)
        except Exception as exc:  # noqa: BLE001 - kullanıcıya hata göstermek için genel yakalama
            self.root.after(0, lambda: self._on_error(exc))
            return

        self.root.after(0, lambda: self._on_success(target_file, last_month))

    def _on_success(self, target_file: Path, last_month: str) -> None:
        self._set_running_state(False)
        self.status.set(f"Tamamlandı. Son ay: {last_month}. Çıktı: {target_file}")
        messagebox.showinfo(
            "Başarılı",
            f"Güncelleme tamamlandı.\n\nSon ay: {last_month}\nÇıktı dosyası:\n{target_file}",
        )

    def _on_error(self, exc: Exception) -> None:
        self._set_running_state(False)
        self.status.set("Hata oluştu. Ayrıntılar için aşağıdaki mesajı okuyun.")
        messagebox.showerror("Hata", f"İşlem başarısız oldu:\n\n{exc}")

    def _set_running_state(self, is_running: bool) -> None:
        state = "disabled" if is_running else "normal"
        self.run_button.config(state=state)
        for child in self.root.winfo_children():
            if isinstance(child, tk.Entry) or isinstance(child, tk.Button) and child is not self.run_button:
                child.config(state=state)
        self.root.update_idletasks()


def main() -> None:
    root = tk.Tk()
    app = KonsolidasyonApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

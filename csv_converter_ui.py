import os
import re
import tkinter as tk
from tkinter import ttk, colorchooser, filedialog, messagebox

from converter_functions import (
    COLOR_PRESETS,
    NAME_TO_HEX,
    PRESET_NAMES,
    configure_hyphenation,
    create_temp_csv_without_empty_columns,
    create_temp_csv_without_selected_columns,
    make_hyphenator,
    normalize_rows,
    read_cochrane_txt,
    read_csv,
    render_pages_dynamic,
    resource_path,
    save_images_as_pdf,
    save_as_excel,
    DPI,
)

class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("CSV → A4 Tabelle (PNG/JPG/PDF/XLSX)")
        self.geometry("900x710")
        self.resizable(False, False)

        self._set_icon()

        self.csv_path = tk.StringVar()
        self.save_path = tk.StringVar()
        self.sep_choice = tk.StringVar(value=",")
        self.sep_custom = tk.StringVar()
        self.format_choice = tk.StringVar(value="PDF")
        self.orientation = tk.StringVar(value="landscape")
        self.use_utf8 = tk.BooleanVar(value=True)

        self.header_color_hex = tk.StringVar(value="#F5F5F5")
        self.zebra_color_hex = tk.StringVar(value="#FCFCFC")
        self.header_color_choice = tk.StringVar(value="Hellgrau")
        self.zebra_color_choice = tk.StringVar(value="Wolkenweiß")
        self.hyphen_choice = tk.StringVar(value="Auto (DE/EN)")
        self.remove_empty_cols = tk.BooleanVar(value=True)  # default AN

        # Kopftext (rechts oben)
        self.header_right_text = tk.StringVar()

        # Schriftgröße
        self.use_custom_font = tk.BooleanVar(value=False)
        self.custom_font_pt = tk.IntVar(value=24)

        # Manuelles Spaltenlöschen: "2;5;7"
        self.remove_cols_str = tk.StringVar(value="")

        self.build_ui()

    def _set_icon(self) -> None:
        try:
            ico_path = resource_path("csvConverter.ico")
            if os.path.exists(ico_path):
                self.iconbitmap(ico_path)
        except Exception:
            pass

    def _make_swatch(self, parent, hex_color: str) -> tk.Canvas:
        canvas = tk.Canvas(parent, width=28, height=18, highlightthickness=1, highlightbackground="#999")
        canvas.create_rectangle(0, 0, 28, 18, fill=hex_color, outline="")
        return canvas

    def _update_swatch(self, canvas: tk.Canvas, hex_color: str) -> None:
        canvas.delete("all")
        canvas.create_rectangle(0, 0, 28, 18, fill=hex_color, outline="")

    def build_ui(self) -> None:
        pad = 8
        frame = ttk.Frame(self)
        frame.pack(fill="both", expand=True, padx=14, pady=14)

        ttk.Label(frame, text="Datei (CSV oder Cochrane TXT):").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.csv_path, width=86).grid(row=0, column=1, sticky="we", padx=(pad, 0))
        ttk.Button(frame, text="Durchsuchen…", command=self.choose_csv).grid(row=0, column=2, padx=(pad, 0))

        ttk.Label(frame, text="Speichern als:").grid(row=1, column=0, sticky="w", pady=(pad, 0))
        ttk.Entry(frame, textvariable=self.save_path, width=86).grid(row=1, column=1, sticky="we", padx=(pad, 0), pady=(pad, 0))
        ttk.Button(frame, text="Ziel wählen…", command=self.choose_save).grid(row=1, column=2, padx=(pad, 0), pady=(pad, 0))

        row2 = ttk.Frame(frame)
        row2.grid(row=2, column=0, columnspan=3, sticky="we", pady=(pad, 0))
        ttk.Label(row2, text="Ausgabeformat:").pack(side="left")
        fmt_box = ttk.Combobox(row2, textvariable=self.format_choice, values=["PNG", "JPG", "PDF", "XLSX"], width=8, state="readonly")
        fmt_box.pack(side="left", padx=(6, 20))
        fmt_box.bind("<<ComboboxSelected>>", lambda _: self.update_save_extension())

        ttk.Label(row2, text="Ausrichtung:").pack(side="left")
        ttk.Radiobutton(row2, text="Hochformat", value="portrait", variable=self.orientation).pack(side="left", padx=4)
        ttk.Radiobutton(row2, text="Querformat", value="landscape", variable=self.orientation).pack(side="left", padx=12)

        colors_row = ttk.Labelframe(frame, text="Farben")
        colors_row.grid(row=3, column=0, columnspan=3, sticky="we", pady=(pad, 0))

        ttk.Label(colors_row, text="Header:").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.header_combo = ttk.Combobox(colors_row, values=PRESET_NAMES, textvariable=self.header_color_choice, width=20, state="readonly")
        self.header_combo.grid(row=0, column=1, sticky="w", padx=(0, 6))
        self.header_combo.bind("<<ComboboxSelected>>", self.on_header_color_change)
        self.header_swatch = self._make_swatch(colors_row, self.header_color_hex.get())
        self.header_swatch.grid(row=0, column=2, padx=(0, 6))
        ttk.Button(colors_row, text="Wählen…", command=lambda: self.pick_custom_color("header")).grid(row=0, column=3, padx=4)

        ttk.Label(colors_row, text="Zebra:").grid(row=0, column=4, sticky="w", padx=(20, 6))
        self.zebra_combo = ttk.Combobox(colors_row, values=PRESET_NAMES, textvariable=self.zebra_color_choice, width=20, state="readonly")
        self.zebra_combo.grid(row=0, column=5, sticky="w", padx=(0, 6))
        self.zebra_combo.bind("<<ComboboxSelected>>", self.on_zebra_color_change)
        self.zebra_swatch = self._make_swatch(colors_row, self.zebra_color_hex.get())
        self.zebra_swatch.grid(row=0, column=6, padx=(0, 6))
        ttk.Button(colors_row, text="Wählen…", command=lambda: self.pick_custom_color("zebra")).grid(row=0, column=7, padx=4)

        row3 = ttk.Frame(frame)
        row3.grid(row=4, column=0, columnspan=3, sticky="we", pady=(pad, 0))
        ttk.Label(row3, text="Trennzeichen:").pack(side="left")
        seps = [",", ";", "Tab", "|", "EBSCO (,)", "PubMed (,)", "Benutzerdefiniert"]
        sep_box = ttk.Combobox(row3, textvariable=self.sep_choice, values=seps, width=16, state="readonly")
        sep_box.pack(side="left", padx=(6, 10))
        sep_box.bind("<<ComboboxSelected>>", self.on_sep_change)
        self.custom_sep_entry = ttk.Entry(row3, textvariable=self.sep_custom, width=8, state="disabled")
        self.custom_sep_entry.pack(side="left")
        ttk.Label(row3, text="(bei 'Benutzerdefiniert' hier Zeichen eingeben)").pack(side="left", padx=(6, 10))

        ttk.Label(frame, text="Hinweis: EBSCO = ','   |   PubMed = ','", foreground="#444").grid(
            row=5, column=0, columnspan=3, sticky="w", pady=(6, 0)
        )
        ttk.Checkbutton(frame, text="UTF-8 korrekt darstellen (ö, ä, ü, ß …)", variable=self.use_utf8).grid(
            row=6, column=0, columnspan=3, sticky="w", pady=(pad, 0)
        )

        ttk.Label(frame, text="Silbentrennung (Header):").grid(row=7, column=0, sticky="w", pady=(pad, 0))
        ttk.Combobox(
            frame,
            textvariable=self.hyphen_choice,
            values=["Auto (DE/EN)", "Deutsch (de_DE)", "Englisch (en_US)", "Aus (keine)"],
            width=18,
            state="readonly"
        ).grid(row=7, column=1, sticky="w", pady=(pad, 0))

        ttk.Checkbutton(
            frame,
            text="Leere Spalten aus CSV entfernen",
            variable=self.remove_empty_cols
        ).grid(row=8, column=0, columnspan=3, sticky="w", pady=(pad, 0))

        # Manuelles Spaltenlöschen (1-basiert, Semikolon-getrennt)
        rm_frame = ttk.Frame(frame)
        rm_frame.grid(row=9, column=0, columnspan=3, sticky="we", pady=(pad, 0))
        ttk.Label(rm_frame, text="Spalten entfernen (1-basiert; Semikolon getrennt):").pack(side="left")
        ttk.Entry(rm_frame, textvariable=self.remove_cols_str, width=40).pack(side="left", padx=(6, 0))

        # Kopftext & Schriftgröße
        head_box = ttk.Labelframe(frame, text="Kopfband (oberhalb der Tabelle) & Schrift")
        head_box.grid(row=10, column=0, columnspan=3, sticky="we", pady=(pad, 0))
        ttk.Label(head_box, text="Text rechts oben:").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(head_box, textvariable=self.header_right_text, width=60).grid(row=0, column=1, sticky="w", padx=(0, 6), pady=6)

        font_row = ttk.Frame(head_box)
        font_row.grid(row=0, column=2, sticky="e", padx=6, pady=6)
        self.font_check = ttk.Checkbutton(font_row, text="Schriftgröße festlegen (pt)", variable=self.use_custom_font, command=self.on_font_toggle)
        self.font_check.pack(side="left", padx=(0, 6))
        self.font_spin = ttk.Spinbox(font_row, from_=8, to=72, width=5, textvariable=self.custom_font_pt, state="disabled")
        self.font_spin.pack(side="left")

        # Button
        ttk.Button(frame, text="Erstellen", command=self.run).grid(row=11, column=2, sticky="e", pady=(pad * 2, 0))
        text = """Keine Gewähr für Richtigkeit, Vollständigkeit und Aktualität.
        Dieses Programm enthält Open-Source-Komponenten:
        - Python (CPython) – PSF License 2.0
        - Pillow – Historical Permission Notice and Disclaimer (HPND)
        - openpyxl – MIT License
        - Pyphen – LGPL-2.1-or-later

        Die jeweiligen Lizenztexte liegen diesem Programm bei.
        Es werden keine Rechte an Marken/Daten Dritter übertragen."""

        style = ttk.Style(self)
        style.configure("Disclaimer.TLabel", foreground="#666", font=("TkDefaultFont", 9))
        ttk.Label(
            frame,
            text=text,
            wraplength=820,
            justify="left"
        ).grid(row=12, column=0, columnspan=3, sticky="w", pady=(6, 0))

        frame.columnconfigure(1, weight=1)

    def on_font_toggle(self) -> None:
        self.font_spin.configure(state="normal" if self.use_custom_font.get() else "disabled")

    def on_header_color_change(self, _event=None) -> None:
        name = self.header_color_choice.get()
        if name == "Benutzerdefiniert…":
            self.pick_custom_color("header")
        else:
            self.header_color_hex.set(NAME_TO_HEX.get(name, self.header_color_hex.get()))
            self._update_swatch(self.header_swatch, self.header_color_hex.get())

    def on_zebra_color_change(self, _event=None) -> None:
        name = self.zebra_color_choice.get()
        if name == "Benutzerdefiniert…":
            self.pick_custom_color("zebra")
        else:
            self.zebra_color_hex.set(NAME_TO_HEX.get(name, self.zebra_color_hex.get()))
            self._update_swatch(self.zebra_swatch, self.zebra_color_hex.get())

    def pick_custom_color(self, target: str) -> None:
        initial = self.header_color_hex.get() if target == "header" else self.zebra_color_hex.get()
        _, hex_color = colorchooser.askcolor(initialcolor=initial, title="Header-Farbe wählen" if target == "header" else "Zebra-Farbe wählen")
        if hex_color:
            if target == "header":
                self.header_color_hex.set(hex_color)
                self.header_color_choice.set("Benutzerdefiniert…")
                self._update_swatch(self.header_swatch, hex_color)
            else:
                self.zebra_color_hex.set(hex_color)
                self.zebra_color_choice.set("Benutzerdefiniert…")
                self._update_swatch(self.zebra_swatch, hex_color)

    def on_sep_change(self, _event=None) -> None:
        if self.sep_choice.get() == "Benutzerdefiniert":
            self.custom_sep_entry.configure(state="normal")
            self.custom_sep_entry.focus_set()
        else:
            self.custom_sep_entry.configure(state="disabled")

    def choose_csv(self) -> None:
        path = filedialog.askopenfilename(
            title="Datei auswählen",
            filetypes=[
                ("CSV oder Cochrane TXT", "*.csv;*.txt"),
                ("CSV", "*.csv"),
                ("TXT", "*.txt"),
                ("Alle Dateien", "*.*"),
            ],
        )
        if path:
            self.csv_path.set(path)

    def update_save_extension(self) -> None:
        current = self.save_path.get().strip()
        if not current:
            return
        base, _ = os.path.splitext(current)
        ext = "." + (self.format_choice.get().lower() if self.format_choice.get() != "JPG" else "jpg")
        self.save_path.set(base + ext)

    def choose_save(self) -> None:
        fmt = self.format_choice.get()
        default_ext = "." + (fmt.lower() if fmt != "JPG" else "jpg")
        filetypes = {
            "PNG": [("PNG-Bild", "*.png")],
            "JPG": [("JPEG-Bild", "*.jpg;*.jpeg")],
            "PDF": [("PDF", "*.pdf")],
            "XLSX": [("Excel-Arbeitsmappe", "*.xlsx")],
        }[fmt]
        path = filedialog.asksaveasfilename(title="Zieldatei wählen", defaultextension=default_ext, filetypes=filetypes)
        if path:
            self.save_path.set(path)

    def get_separator(self) -> str:
        choice = self.sep_choice.get()
        if choice == "Tab":
            return "\t"
        if choice in ("EBSCO (,)", "PubMed (,)"):
            return ","
        if choice == "Benutzerdefiniert":
            custom = self.sep_custom.get()
            if not custom:
                messagebox.showerror("Fehler", "Bitte benutzerdefiniertes Trennzeichen eingeben.")
                raise ValueError
            if len(custom) != 1:
                messagebox.showwarning("Hinweis", "Es wird nur das erste Zeichen als Trennzeichen genutzt.")
            return custom[0]
        return choice

    def parse_remove_cols_spec(self, spec: str) -> list[int]:
        if not spec or not spec.strip():
            return []
        tokens = re.split(r"[;]+", spec.strip())
        result: set[int] = set()
        for t in tokens:
            t = t.strip()
            if not t:
                continue
            try:
                n = int(t)
                if n >= 1:
                    result.add(n)
            except Exception:
                pass
        return sorted(result)

    def run(self) -> None:
        if not self.csv_path.get():
            messagebox.showerror("Fehler", "Bitte eine CSV- oder TXT-Datei auswählen.")
            return
        if not self.save_path.get():
            self.choose_save()
            if not self.save_path.get():
                return

        temp_paths = []  # temporäre Dateien
        try:
            input_path = self.csv_path.get().strip()
            header: list[str] = []
            data: list[list[str]] = []
            meta = {}

            if input_path.lower().endswith(".txt"):
                header, data, meta = read_cochrane_txt(input_path)
            else:
                separator = self.get_separator()
                input_for_reading = input_path

                # Preview (Spaltenzahl)
                try:
                    rows_preview = normalize_rows(read_csv(input_for_reading, separator, encoding_utf8=self.use_utf8.get()))
                    if not rows_preview:
                        messagebox.showerror("Fehler", "Die CSV-Datei enthält keine verwertbaren Daten.")
                        return
                    max_cols_preview = max(len(r) for r in rows_preview)
                except Exception as exc:
                    messagebox.showerror("Fehler", f"CSV konnte nicht gelesen werden:\n{exc}")
                    return

                # Manuelles Spaltenlöschen
                try:
                    cols_spec = self.remove_cols_str.get().strip()
                    cols_to_remove = self.parse_remove_cols_spec(cols_spec)
                    if cols_to_remove:
                        invalid = [n for n in cols_to_remove if n < 1 or n > max_cols_preview]
                        if invalid:
                            messagebox.showerror(
                                "Fehler",
                                f"Ungültige Spaltennummer(n): {invalid}\n"
                                f"Die Datei enthält nur {max_cols_preview} Spalten."
                            )
                            return
                        tmp_path = create_temp_csv_without_selected_columns(
                            input_for_reading,
                            separator,
                            columns_to_remove_1based=cols_to_remove,
                            encoding_utf8=self.use_utf8.get(),
                        )
                        temp_paths.append(tmp_path)
                        input_for_reading = tmp_path
                except Exception as exc:
                    messagebox.showerror("Fehler", f"Spaltenentfernung fehlgeschlagen:\n{exc}")
                    return

                # Leere Spalten entfernen (optional)
                if self.remove_empty_cols.get():
                    try:
                        tmp_path = create_temp_csv_without_empty_columns(
                            input_for_reading,
                            separator,
                            encoding_utf8=self.use_utf8.get(),
                        )
                        temp_paths.append(tmp_path)
                        input_for_reading = tmp_path
                    except Exception as exc:
                        messagebox.showerror("Fehler", f"Spaltenbereinigung (leer) fehlgeschlagen:\n{exc}")
                        return

                rows = normalize_rows(read_csv(input_for_reading, separator, encoding_utf8=self.use_utf8.get()))
                if not rows:
                    messagebox.showerror("Fehler", "Die CSV-Datei enthält keine verwertbaren Daten.")
                    return
                header, data = rows[0], rows[1:]
                if not isinstance(header, (list, tuple)):
                    header = [str(header)]

            top_note = f"Date Run: {meta['Date Run']}" if meta.get("Date Run") else None

            # Silbentrennung konfigurieren
            choice = self.hyphen_choice.get()
            hyphenator = make_hyphenator(choice, header)
            configure_hyphenation(enable_headers=bool(hyphenator) and not choice.startswith("Aus"), hyphenator=hyphenator)

            fmt = self.format_choice.get()
            orientation = self.orientation.get()
            out_path = self.save_path.get()

            right_text = self.header_right_text.get().strip() or None
            custom_pt = int(self.custom_font_pt.get()) if self.use_custom_font.get() else None

            if fmt in ("PNG", "JPG", "PDF"):
                # 1) Seiten rendern
                pages = render_pages_dynamic(
                    header, data, orientation,
                    self.header_color_hex.get(), self.zebra_color_hex.get(),
                    top_note=top_note,
                    top_right_text=right_text,
                    custom_font_pt=custom_pt
                )

                # 2) Export
                if fmt == "PDF":
                    save_images_as_pdf(pages, out_path)
                else:
                    base, ext = os.path.splitext(out_path)
                    if len(pages) == 1:
                        image = pages[0]
                        if fmt == "PNG":
                            image.save(out_path, "PNG", dpi=(DPI, DPI))
                        else:
                            image.convert("RGB").save(out_path, "JPEG", quality=95, dpi=(DPI, DPI))
                    else:
                        for idx, image in enumerate(pages, start=1):
                            filename = f"{base}_{idx:02d}{ext if fmt == 'PNG' else '.jpg'}"
                            if fmt == "PNG":
                                image.save(filename, "PNG", dpi=(DPI, DPI))
                            else:
                                image.convert("RGB").save(filename, "JPEG", quality=95, dpi=(DPI, DPI))
                        out_path = f"{base}_01{ext if fmt == 'PNG' else '.jpg'}"
            else:
                # XLSX
                save_as_excel(header, data, out_path, orientation, fit_width=True, meta_note=top_note, header_right_text=right_text)

            messagebox.showinfo("Fertig", f"Erfolgreich gespeichert:\n{out_path}")

        except Exception as exc:
            messagebox.showerror("Fehler", f"Beim Erstellen ist ein Fehler aufgetreten:\n{exc}")
        finally:
            try:
                for p in temp_paths:
                    if p and os.path.exists(p):
                        os.remove(p)
            except Exception:
                pass
            configure_hyphenation(False, None)

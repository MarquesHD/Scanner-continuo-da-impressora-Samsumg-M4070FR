"""
Scanner Samsung M4070FR -> PDF
--------------------------------
App desktop para Windows que digitaliza documentos usando o driver
WIA (Windows Image Acquisition) da impressora Samsung M4070FR e
salva sempre em PDF, com opção de digitalização contínua (várias
páginas no mesmo PDF), nome de arquivo editável e exclusão de
páginas individuais da sessão.

Layout inspirado em telas de onboarding modernas: painel lateral
azul com indicador de progresso + painel branco com o formulário.

Requisitos (ver requirements.txt):
    pip install customtkinter pywin32 img2pdf Pillow
"""

import os
import sys
import tempfile
import threading
from datetime import datetime

import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image
import img2pdf

try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None
    pythoncom = None

# GUID padrão do formato BMP no WIA (constante fixa da API do Windows)
WIA_FORMAT_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
WIA_DEVICE_TYPE_SCANNER = 1

# ---------------------------------------------------------------- tema ----
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

NAVY = "#1a3fc4"          # painel lateral
ACCENT = "#2f5bff"
ACCENT_HOVER = "#2649d6"
STEP_MUTED = "#8fa3f0"
SUCCESS = "#22a06b"
SUCCESS_HOVER = "#1c8759"
DANGER = "#e5484d"
DANGER_HOVER = "#c93d41"
CARD_RADIUS = 16
SIDEBAR_W = 250


class ScannerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Scanner Samsung M4070FR")
        self.geometry("820x720")
        self.minsize(720, 620)
        self.configure(fg_color="#f5f6fa")

        self.output_folder = os.path.join(os.path.expanduser("~"), "Documents", "Digitalizados")
        os.makedirs(self.output_folder, exist_ok=True)

        self.session_pages = []       # caminhos temporários das páginas (bmp)
        self.thumbnail_widgets = []   # widgets de miniatura na lista
        self.step_rows = []           # widgets do indicador de progresso
        self.scanning = False

        self._build_layout()
        self._update_steps()

    # ------------------------------------------------------- layout ----
    def _build_layout(self):
        root = ctk.CTkFrame(self, fg_color="transparent")
        root.pack(fill="both", expand=True)

        self.sidebar = ctk.CTkFrame(root, width=SIDEBAR_W, corner_radius=0, fg_color=NAVY)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        content_wrap = ctk.CTkFrame(root, fg_color="#f5f6fa")
        content_wrap.pack(side="left", fill="both", expand=True)

        self.content = ctk.CTkScrollableFrame(content_wrap, fg_color="#f5f6fa")
        self.content.pack(fill="both", expand=True, padx=36, pady=32)

        self._build_sidebar()
        self._build_form()
        self._build_pages_section()

    # ------------------------------------------------------- sidebar ----
    def _build_sidebar(self):
        top = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        top.pack(fill="x", padx=28, pady=(32, 8))

        badge = ctk.CTkFrame(top, width=42, height=42, corner_radius=12, fg_color="#ffffff")
        badge.pack(anchor="w")
        badge.pack_propagate(False)
        ctk.CTkLabel(badge, text="🖨️", font=ctk.CTkFont(size=20)).pack(expand=True)

        ctk.CTkLabel(
            self.sidebar, text="Digitalizar seus\ndocumentos",
            font=ctk.CTkFont(size=20, weight="bold"), text_color="#ffffff",
            justify="left", anchor="w"
        ).pack(fill="x", padx=28, pady=(20, 4))

        ctk.CTkLabel(
            self.sidebar, text="Samsung M4070FR",
            font=ctk.CTkFont(size=13), text_color=STEP_MUTED, anchor="w"
        ).pack(fill="x", padx=28, pady=(0, 26))

        self.steps_container = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.steps_container.pack(fill="x", padx=28)

        # rodapé com resumo da sessão
        self.sidebar_footer = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.sidebar_footer.pack(side="bottom", fill="x", padx=28, pady=28)

        ctk.CTkFrame(self.sidebar_footer, height=1, fg_color="#3f5cd4").pack(fill="x", pady=(0, 14))

        self.footer_filename = ctk.CTkLabel(
            self.sidebar_footer, text="", font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#ffffff", anchor="w", justify="left", wraplength=190
        )
        self.footer_filename.pack(fill="x")
        self.footer_pages = ctk.CTkLabel(
            self.sidebar_footer, text="", font=ctk.CTkFont(size=12),
            text_color=STEP_MUTED, anchor="w"
        )
        self.footer_pages.pack(fill="x", pady=(2, 0))

    def _update_steps(self):
        for w in self.step_rows:
            w.destroy()
        self.step_rows = []

        stage = "scanning" if self.session_pages else "config"
        steps = [
            (1, "Configurar arquivo", "done"),
            (2, "Digitalizar páginas", "current" if stage == "scanning" else "pending"),
            (3, "Salvar em PDF", "pending"),
        ]

        for i, (num, label, state) in enumerate(steps):
            row = ctk.CTkFrame(self.steps_container, fg_color="transparent")
            row.pack(fill="x")

            circle = ctk.CTkFrame(row, width=30, height=30, corner_radius=15)
            circle.pack(side="left")
            circle.pack_propagate(False)

            if state == "done":
                circle.configure(fg_color="#ffffff")
                ctk.CTkLabel(circle, text="✓", font=ctk.CTkFont(size=13, weight="bold"),
                             text_color=NAVY).pack(expand=True)
                text_color = "#ffffff"
            elif state == "current":
                circle.configure(fg_color=NAVY, border_width=2, border_color="#ffffff")
                ctk.CTkLabel(circle, text=str(num), font=ctk.CTkFont(size=13, weight="bold"),
                             text_color="#ffffff").pack(expand=True)
                text_color = "#ffffff"
            else:
                circle.configure(fg_color="transparent", border_width=1, border_color=STEP_MUTED)
                ctk.CTkLabel(circle, text=str(num), font=ctk.CTkFont(size=13),
                             text_color=STEP_MUTED).pack(expand=True)
                text_color = STEP_MUTED

            ctk.CTkLabel(row, text=label, font=ctk.CTkFont(size=13, weight="bold" if state != "pending" else "normal"),
                         text_color=text_color).pack(side="left", padx=12)

            self.step_rows.append(row)

            if i < len(steps) - 1:
                connector = ctk.CTkFrame(self.steps_container, width=2, height=22,
                                          fg_color="#3f5cd4")
                connector.pack(padx=14, anchor="w")
                self.step_rows.append(connector)

        # atualiza rodapé
        filename = self.filename_entry.get().strip() if hasattr(self, "filename_entry") else self._default_filename()
        self.footer_filename.configure(text=f"📄 {filename}.pdf" if filename else "📄 sem nome")
        self.footer_pages.configure(text=f"{len(self.session_pages)} página(s) na sessão atual")

    # --------------------------------------------------------- form ----
    def _build_form(self):
        ctk.CTkLabel(
            self.content, text="Farmacia Joao Paulo II",
            font=ctk.CTkFont(size=26, weight="bold"), anchor="w"
        ).pack(fill="x")
        ctk.CTkLabel(
            self.content, text="Feito por G.Marques, para uso interno.",
            font=ctk.CTkFont(size=9), text_color="gray", anchor="w"
        ).pack(fill="x", pady=(2, 22))

        card = ctk.CTkFrame(self.content, corner_radius=CARD_RADIUS, fg_color="#ffffff",
                             border_width=1, border_color="#e7e9f2")
        card.pack(fill="x")

        ctk.CTkLabel(card, text="Nome do arquivo", font=ctk.CTkFont(size=12, weight="bold"),
                     text_color="#555", anchor="w").pack(fill="x", padx=22, pady=(20, 6))
        name_row = ctk.CTkFrame(card, fg_color="transparent")
        name_row.pack(fill="x", padx=22)
        self.filename_entry = ctk.CTkEntry(
            name_row, placeholder_text="documento", height=42, corner_radius=10,
            border_width=1, border_color="#dde1ef", fg_color="#f8f9fc"
        )
        self.filename_entry.insert(0, self._default_filename())
        self.filename_entry.bind("<KeyRelease>", lambda e: self._update_steps())
        self.filename_entry.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(name_row, text=".pdf", font=ctk.CTkFont(size=13, weight="bold"),
                     text_color="gray").pack(side="left", padx=(10, 0))

        ctk.CTkLabel(card, text="Salvar em", font=ctk.CTkFont(size=12, weight="bold"),
                     text_color="#555", anchor="w").pack(fill="x", padx=22, pady=(18, 6))
        folder_row = ctk.CTkFrame(card, height=42, corner_radius=10, fg_color="#f8f9fc",
                                   border_width=1, border_color="#dde1ef")
        folder_row.pack(fill="x", padx=22)
        folder_row.pack_propagate(False)
        self.folder_label = ctk.CTkLabel(
            folder_row, text=self._short_path(self.output_folder),
            text_color="#333", anchor="w", font=ctk.CTkFont(size=13)
        )
        self.folder_label.pack(side="left", fill="x", expand=True, padx=14)
        ctk.CTkButton(
            folder_row, text="Trocar", width=70, height=30, corner_radius=8,
            fg_color=ACCENT, hover_color=ACCENT_HOVER, command=self.choose_folder
        ).pack(side="right", padx=6)

        toggle_row = ctk.CTkFrame(card, fg_color="transparent")
        toggle_row.pack(fill="x", padx=22, pady=(20, 20))
        col = ctk.CTkFrame(toggle_row, fg_color="transparent")
        col.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(col, text="Digitalização contínua", font=ctk.CTkFont(size=13, weight="bold"),
                     anchor="w").pack(fill="x")
        ctk.CTkLabel(col, text="Junta várias páginas em um único PDF",
                     font=ctk.CTkFont(size=12), text_color="gray", anchor="w").pack(fill="x")
        self.continuous_var = ctk.BooleanVar(value=True)
        ctk.CTkSwitch(toggle_row, text="", variable=self.continuous_var,
                       progress_color=ACCENT, button_color="#ffffff").pack(side="right")

        self.scan_btn = ctk.CTkButton(
            self.content, text="🖨️   Digitalizar página", height=50, corner_radius=13,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color=ACCENT, hover_color=ACCENT_HOVER, command=self.start_scan
        )
        self.scan_btn.pack(fill="x", pady=(20, 10))

        actions_row = ctk.CTkFrame(self.content, fg_color="transparent")
        actions_row.pack(fill="x")
        self.finish_btn = ctk.CTkButton(
            actions_row, text="✅  Finalizar e salvar PDF", height=42, corner_radius=12,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=SUCCESS, hover_color=SUCCESS_HOVER, command=self.finish_pdf
        )
        self.finish_btn.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.cancel_btn = ctk.CTkButton(
            actions_row, text="🗑️  Cancelar sessão", height=42, corner_radius=12,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="transparent", hover_color="#fbe4e4", text_color=DANGER,
            border_width=1, border_color=DANGER, command=self.cancel_session
        )
        self.cancel_btn.pack(side="right", fill="x", expand=True, padx=(6, 0))

        self.status_label = ctk.CTkLabel(self.content, text="Pronto para digitalizar", text_color="gray")
        self.status_label.pack(pady=(12, 0))

    # ---------------------------------------------------- pages list ----
    def _build_pages_section(self):
        header_row = ctk.CTkFrame(self.content, fg_color="transparent")
        header_row.pack(fill="x", pady=(26, 8))
        self.pages_count_label = ctk.CTkLabel(
            header_row, text="Páginas digitalizadas: 0",
            font=ctk.CTkFont(size=14, weight="bold"), anchor="w"
        )
        self.pages_count_label.pack(side="left")

        self.pages_frame = ctk.CTkFrame(self.content, corner_radius=CARD_RADIUS,
                                         fg_color="#eef0f8")
        self.pages_frame.pack(fill="x")

        self._empty_state_label = ctk.CTkLabel(
            self.pages_frame, text="Nenhuma página digitalizada ainda",
            text_color="gray", font=ctk.CTkFont(size=13)
        )
        self._empty_state_label.pack(pady=36)

    # -------------------------------------------------------- utils ----
    def _default_filename(self):
        return f"digitalizacao_{datetime.now().strftime('%Y%m%d_%H%M')}"

    def _short_path(self, path):
        return path if len(path) < 40 else "..." + path[-37:]

    def choose_folder(self):
        folder = filedialog.askdirectory(initialdir=self.output_folder)
        if folder:
            self.output_folder = folder
            self.folder_label.configure(text=self._short_path(folder))

    # -------------------------------------------------------- scan -----
    def start_scan(self):
        if self.scanning:
            return
        self.scanning = True
        self.scan_btn.configure(state="disabled", text="Digitalizando...")
        self.status_label.configure(text="Comunicando com a impressora...")
        threading.Thread(target=self._scan_worker, daemon=True).start()

    def _scan_worker(self):
        # Toda thread que usa COM (win32com) precisa inicializar o COM
        # nela mesma antes de qualquer chamada, senão dá o erro
        # "CoInitialize não foi chamado".
        if pythoncom is not None:
            pythoncom.CoInitialize()
        try:
            image_path = self._perform_wia_scan()
            self.after(0, self._on_scan_success, image_path)
        except Exception as e:
            self.after(0, self._on_scan_error, str(e))
        finally:
            if pythoncom is not None:
                pythoncom.CoUninitialize()

    def _perform_wia_scan(self):
        if win32com is None:
            raise RuntimeError(
                "A biblioteca pywin32 não está instalada.\n"
                "Rode: pip install pywin32"
            )

        device_manager = win32com.client.Dispatch("WIA.DeviceManager")
        scanner_info = None
        for info in device_manager.DeviceInfos:
            if info.Type == WIA_DEVICE_TYPE_SCANNER:
                scanner_info = info
                break

        if scanner_info is None:
            raise RuntimeError(
                "Nenhum scanner encontrado.\n\n"
                "Verifique se:\n"
                "- A Samsung M4070FR está ligada e conectada (USB ou rede)\n"
                "- O driver de scanner (WIA) está instalado no Windows\n"
                "- Ela aparece em 'Digitalizadores e câmeras' no Painel de Controle"
            )

        device = scanner_info.Connect()
        item = device.Items[1]

        image = item.Transfer(WIA_FORMAT_BMP)

        tmp_path = os.path.join(
            tempfile.gettempdir(),
            f"scan_page_{len(self.session_pages)}_{datetime.now().strftime('%H%M%S%f')}.bmp"
        )
        image.SaveFile(tmp_path)
        return tmp_path

    def _on_scan_success(self, image_path):
        self.session_pages.append(image_path)
        self._refresh_pages_list()
        self._update_steps()
        self.scanning = False
        self.scan_btn.configure(state="normal", text="🖨️   Digitalizar página")

        if self.continuous_var.get():
            self.status_label.configure(text="Coloque a próxima folha e clique em Digitalizar")
        else:
            self.status_label.configure(text="Página digitalizada com sucesso")

    def _on_scan_error(self, error_msg):
        self.scanning = False
        self.scan_btn.configure(state="normal", text="🖨️   Digitalizar página")
        self.status_label.configure(text="Erro ao digitalizar")
        messagebox.showerror("Erro na digitalização", error_msg)

    # -------------------------------------------------------- pages ----
    def _refresh_pages_list(self):
        for w in self.thumbnail_widgets:
            w.destroy()
        self.thumbnail_widgets = []

        if not self.session_pages:
            self._empty_state_label.pack(pady=36)
        else:
            self._empty_state_label.pack_forget()
            for idx, path in enumerate(self.session_pages, start=1):
                self._add_thumbnail(path, idx)

        self.pages_count_label.configure(text=f"Páginas digitalizadas: {len(self.session_pages)}")

    def _add_thumbnail(self, image_path, page_num):
        try:
            img = Image.open(image_path)
            img.thumbnail((52, 68))
            ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)
        except Exception:
            ctk_img = None

        row = ctk.CTkFrame(self.pages_frame, corner_radius=12, fg_color="#ffffff",
                            border_width=1, border_color="#e2e5f1")
        row.pack(fill="x", pady=6, padx=10)

        left = ctk.CTkFrame(row, fg_color="transparent")
        left.pack(side="left", padx=10, pady=10)
        if ctk_img is not None:
            ctk.CTkLabel(left, image=ctk_img, text="").pack()

        ctk.CTkLabel(row, text=f"Página {page_num}",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(side="left", padx=6)

        ctk.CTkButton(
            row, text="🗑️", width=34, height=34, corner_radius=10,
            fg_color="transparent", hover_color="#fbe4e4",
            text_color=DANGER,
            command=lambda p=image_path: self.delete_page(p)
        ).pack(side="right", padx=10)

        self.thumbnail_widgets.append(row)

    def delete_page(self, image_path):
        if not messagebox.askyesno("Excluir página", "Remover esta página da sessão atual?"):
            return
        if image_path in self.session_pages:
            self.session_pages.remove(image_path)
        for p in (image_path, image_path.replace(".bmp", ".jpg")):
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass
        self._refresh_pages_list()
        self._update_steps()
        if not self.session_pages:
            self.status_label.configure(text="Pronto para digitalizar")

    # -------------------------------------------------------- save -----
    def finish_pdf(self):
        if not self.session_pages:
            messagebox.showwarning("Nenhuma página", "Digitalizar.")
            return

        filename = self.filename_entry.get().strip()
        if not filename:
            messagebox.showwarning("Nome inválido", "Digite um nome para o arquivo.")
            return
        if not filename.lower().endswith(".pdf"):
            filename += ".pdf"

        output_path = os.path.join(self.output_folder, filename)

        if os.path.exists(output_path):
            if not messagebox.askyesno("Arquivo já existe", f"'{filename}' já existe. Substituir?"):
                return

        try:
            converted = []
            for p in self.session_pages:
                img = Image.open(p).convert("RGB")
                jpg_path = p.replace(".bmp", ".jpg")
                img.save(jpg_path, "JPEG", quality=92)
                converted.append(jpg_path)

            with open(output_path, "wb") as f:
                f.write(img2pdf.convert(converted))

            messagebox.showinfo("PDF salvo", f"PDF salvo com sucesso em:\n{output_path}")
            self._reset_session()
        except Exception as e:
            messagebox.showerror("Erro ao salvar PDF", str(e))

    def cancel_session(self):
        if not self.session_pages:
            return
        if messagebox.askyesno("Cancelar sessão", "Descartar todas as páginas digitalizadas até agora?"):
            self._reset_session()

    def _reset_session(self):
        for p in self.session_pages:
            for path in (p, p.replace(".bmp", ".jpg")):
                try:
                    if os.path.exists(path):
                        os.remove(path)
                except Exception:
                    pass

        self.session_pages = []
        self._refresh_pages_list()
        self._update_steps()

        self.status_label.configure(text="Pronto para digitalizar")
        self.filename_entry.delete(0, "end")
        self.filename_entry.insert(0, self._default_filename())


if __name__ == "__main__":
    if sys.platform != "win32":
        print("Este aplicativo usa a API WIA do Windows e só funciona no Windows.")
    app = ScannerApp()
    app.mainloop()

import contextlib
from pathlib import Path
from tempfile import NamedTemporaryFile
import sys
import os

import streamlit as st


def get_base_path():
    """Get the base path - works for both frozen and normal mode."""
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # PyInstaller frozen executable
        return Path(sys._MEIPASS)
    else:
        # Normal Python execution
        return Path(__file__).resolve().parent.parent


# Add root to path for imports
ROOT_DIR = get_base_path()
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from update_konsolidasyon import run_update  # noqa: E402  # pylint: disable=wrong-import-position


st.set_page_config(page_title="Konsolidasyon Güncelleme Aracı", page_icon="🧮")
st.title("Konsolidasyon Güncelleme Aracı")
st.write(
    """
    *Data.xlsx* dosyasındaki gerçekleşen veriler ile konsolidasyon raporunu otomatik olarak
    güncelleyin. Gerekli dosyaları yükleyin, ardından **Güncelle** butonuna tıklayın.
    İşlem tamamlandığında güncellenmiş Excel dosyasını indirebilirsiniz.
    """
)


def _persist_upload(uploaded_file):
    """Streamlit dosyasını geçici olarak diske kaydet."""
    tmp = NamedTemporaryFile(delete=False, suffix=".xlsx")
    try:
        tmp.write(uploaded_file.getbuffer())
        tmp.flush()
    finally:
        tmp.close()
    return Path(tmp.name)


def _cleanup_temp_files(*paths):
    for path in paths:
        if path and path.exists():
            with contextlib.suppress(Exception):
                path.unlink()


st.subheader("1. Dosyaları Yükleyin")
uploaded_data_file = st.file_uploader(
    "Data dosyası (.xlsx)", help="Gerçekleşen verilerin bulunduğu data dosyası"
)
uploaded_konsolidasyon_file = st.file_uploader(
    "Konsolidasyon dosyası (.xlsx)",
    help="Güncellenecek konsolidasyon şablonu",
)

st.divider()

st.subheader("2. Güncellemeyi Başlatın")
trigger = st.button("Raporu Güncelle", type="primary")

if trigger:
    def _build_output_name(path_obj):
        return f"{path_obj.stem}_guncel{path_obj.suffix or '.xlsx'}"

    if not uploaded_data_file or not uploaded_konsolidasyon_file:
        st.error("Lütfen her iki Excel dosyasını da yükleyin.")
        st.stop()

    data_path = _persist_upload(uploaded_data_file)
    konsolidasyon_path = _persist_upload(uploaded_konsolidasyon_file)
    cleanup_candidates = [data_path, konsolidasyon_path]
    kons_base = Path(uploaded_konsolidasyon_file.name)
    output_name = _build_output_name(kons_base)

    with NamedTemporaryFile(delete=False, suffix=".xlsx") as output_tmp:
        output_tmp.close()
        cleanup_candidates.append(Path(output_tmp.name))

        try:
            result_path, last_month = run_update(
                data_path, konsolidasyon_path, output_tmp.name
            )
            with open(result_path, "rb") as result_file:
                result_bytes = result_file.read()
        except Exception as exc:
            st.error(f"Güncelleme sırasında hata oluştu: {exc}")
            _cleanup_temp_files(*cleanup_candidates)
            st.stop()

    st.success(f"Rapor başarıyla güncellendi. Son güncellenen ay: **{last_month}**")
    st.download_button(
        "Güncellenmiş dosyayı indir",
        data=result_bytes,
        file_name=output_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    _cleanup_temp_files(*cleanup_candidates)

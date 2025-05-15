###############################################################
#         Reverse-engineering / Data-science Jupyter          #
#  radare2 + r2dec, pwntools, angr, Qiling (all rootfs)       #
###############################################################

# ── Base ─────────────────────────────────────────────────────
FROM python:3.12-slim

# ── 0. Debian backports (нужен свежий radare2) ───────────────
RUN echo "deb http://deb.debian.org/debian bookworm-backports main" \
      > /etc/apt/sources.list.d/backports.list

# ── 1. System toolchain & CLI-утилиты ────────────────────────
RUN apt-get update && \
    DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
        build-essential gcc g++ make cmake pkg-config \
        mingw-w64 gcc-mingw-w64 g++-mingw-w64 binutils-multiarch binutils-mingw-w64 \
        binutils patchelf elfutils file \
        gdb gdb-multiarch strace ltrace \
        vim-common hexedit \
        tshark tcpdump \
        nodejs npm \
        meson ninja-build \
        chromium chromium-driver \
        fonts-liberation libfontconfig1 libnss3 libx11-6 \
        git curl wget libffi-dev libssl-dev libz3-dev && \
    # radare2 + dev backports
    apt-get -y -t bookworm-backports install radare2 libradare2-dev && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

# ── 1.1. r2pm  ───────────────────────────────────────────
RUN r2pm -U

# ── 1.2. r2dec (bug r2pm) ──────────
RUN git clone --depth 1 https://github.com/wargio/r2dec-js /opt/r2dec && \
    cd /opt/r2dec && \
    meson setup build && \
    ninja -C build && \
    ninja -C build install && \
    cd / && rm -rf /opt/r2dec

# ── 1.3. Qiling rootfs (full) ────────────────────────
RUN git clone --depth 1 https://github.com/qilingframework/rootfs.git /opt/qiling/rootfs
ENV QILING_ROOTFS=/opt/qiling/rootfs

# ── 2. Python ───────────────────────────────────────────
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir \
        # Jupyter
        jupyterlab ipykernel notebook \
        # Reverse-engineering / emu
        angr z3-solver qiling \
        yara-python python-evtx \
        capstone unicorn keystone-engine lief pefile dnfile binwalk \
        r2pipe volatility3 pwntools \
        # Network
        pyshark scapy \
        # Data science / visual
        matplotlib pandas seaborn \
        # Web-scraping / auto
        beautifulsoup4 selenium

# ── 3. Jupyter Lab ───────────────────────────────────────────
ENV JUPYTER_TOKEN=123
EXPOSE 8888

CMD jupyter lab --ip=0.0.0.0 --port=8888 --no-browser --allow-root \
    --ServerApp.token="$JUPYTER_TOKEN"

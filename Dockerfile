###############################################################
#    Reverse-engineering Jupyter image (PE + radare2)         #
###############################################################
FROM python:3.12-slim

RUN echo "deb http://deb.debian.org/debian bookworm-backports main" \
        > /etc/apt/sources.list.d/backports.list

RUN apt-get update && \
    DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
        build-essential gcc g++ make cmake pkg-config \
        mingw-w64 gcc-mingw-w64 g++-mingw-w64 binutils-mingw-w64 \
        binutils binutils-multiarch patchelf elfutils file \
        gdb gdb-multiarch strace ltrace \
        vim-common hexedit \
        git curl wget libffi-dev libssl-dev libz3-dev \
    && apt-get -y -t bookworm-backports install radare2 \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

RUN pip install --no-cache-dir \
        jupyterlab ipykernel notebook \
        angr z3-solver \
        yara-python python-evtx \
        capstone unicorn keystone-engine \
        lief pefile dnfile binwalk \
        r2pipe volatility3 pwntools

ENV JUPYTER_TOKEN=123
EXPOSE 8888

CMD jupyter lab --ip=0.0.0.0 --port=8888 --no-browser --allow-root \
    --ServerApp.token="$JUPYTER_TOKEN"

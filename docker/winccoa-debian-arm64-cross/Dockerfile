# WinCC OA for Debian 12 arm64 in an x86_64 Debian 12 image, for building
# CtrlExcelReader for arm64 on an x86_64 host: the arm64 WinCC OA packages
# (unpacked, not installed), a cross compiler and qemu-user, which runs the
# arm64 tests (Rust, C++ and WCCOActrl) on the build host without binfmt.
# The build-and-test image (docker/Dockerfile) goes on top. Build it with the
# arm64 package directory as context:
#
#   docker build -t winccoa321-arm64-cross:3.21.7 -f docker/winccoa-debian-arm64-cross/Dockerfile path/to/WinCCOA-3.21.7-debian.arm64v8
#
# Debian 12 like the packages: the cross compiler builds against its glibc
# and libstdc++, which is what the built .so then requires. For x86_64 hosts
# only: Debian has no aarch64 cross compiler for arm64 itself. On an arm64
# host, build Siemens' Dockerfile natively and docker/Dockerfile on top.
ARG BASE_IMAGE=debian:bookworm-slim
FROM ${BASE_IMAGE}

ARG USER=winccoa
ARG USER_ID=1000
ARG USER_GID=1000

# The arm64 runtime libraries come from Debian's multiarch packages (they
# install next to the x86_64 ones): libc6 and libstdc++6 for the tests, and
# the libraries the Base and BaseServer packages depend on for WCCOActrl.
RUN dpkg --add-architecture arm64 \
 && apt-get update \
 && apt-get install -y --no-install-recommends \
      sudo unzip locales locales-all tzdata \
      g++-aarch64-linux-gnu qemu-user \
      libc6:arm64 libstdc++6:arm64 libgcc-s1:arm64 \
      libbrotli1:arm64 libcurl4:arm64 libdbus-1-3:arm64 libglib2.0-0:arm64 \
      libicu72:arm64 libpam0g:arm64 libpcre2-16-0:arm64 libssl3:arm64 \
      libzstd1:arm64 zlib1g:arm64 libncurses6:arm64 libtinfo6:arm64 \
      libxml2:arm64 libzmq5:arm64 libsdl2-2.0-0:arm64 \
 && apt-get clean && rm -rf /var/lib/apt/lists/*

ENV LANG=en_US.UTF-8
ENV LANGUAGE=en_US:en
ENV LC_ALL=en_US.UTF-8

RUN groupadd --gid "$USER_GID" "$USER" \
 && useradd --uid "$USER_ID" --gid "$USER_GID" -m --home-dir "/home/$USER" "$USER" \
 && echo "$USER ALL=(ALL) NOPASSWD:ALL" >> /etc/sudoers

# Base provides the libraries, BaseServer WCCOActrl, APIcpp the API
# (api/CMakeDefines.txt), CodeMeter only libwibucm (the CodeMeter service is
# not needed: WCCOActrl -n runs without a license). The packages cannot be
# installed with apt on x86_64 (their dependencies would resolve to the arm64
# versions of tools like unzip), so they are unpacked and their postinst
# scripts (shell only) run: Base's registers the installation, BaseServer's
# creates the directories WCCOActrl uses (e.g. log/).
COPY codemeter-lite_*_arm64.deb WinCCOA-*-Base-debian.arm64.deb WinCCOA-*-BaseServer-debian.arm64.deb WinCCOA-*-APIcpp-debian.arm64.deb /tmp/deb/
RUN set -e; \
    mkdir /tmp/codemeter; \
    dpkg-deb -x /tmp/deb/codemeter-lite_*_arm64.deb /tmp/codemeter; \
    cp /tmp/codemeter/usr/lib/aarch64-linux-gnu/libwibucm.so /usr/lib/aarch64-linux-gnu/; \
    for deb in /tmp/deb/WinCCOA-*.deb; do \
      dpkg-deb -x "$deb" /; \
      dpkg-deb -e "$deb" /tmp/control; \
      if [ -f /tmp/control/postinst ]; then bash /tmp/control/postinst configure; fi; \
      rm -rf /tmp/control; \
    done; \
    rm -rf /tmp/deb /tmp/codemeter

# For docker/Dockerfile and run-tests.sh: cross-compile with this toolchain
# file from the sources' cmake/ and this Rust target, run the arm64 binaries
# with qemu.
ENV CTRLEXCEL_TOOLCHAIN=toolchain-aarch64-linux-gnu.cmake
ENV CTRLEXCEL_RUST_TARGET=aarch64-unknown-linux-gnu
ENV CARGO_TARGET_AARCH64_UNKNOWN_LINUX_GNU_LINKER=aarch64-linux-gnu-gcc
ENV CARGO_TARGET_AARCH64_UNKNOWN_LINUX_GNU_RUNNER=qemu-aarch64
ENV CTRLEXCEL_CTRL_RUNNER=qemu-aarch64

USER $USER
ENV OAINST=/opt/WinCC_OA/3.21

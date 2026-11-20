# Build and test image for CtrlExcelReader on Linux: the WinCC OA image
# (built from Siemens' Dockerfile in the WinCC OA Debian package directory)
# plus a C++ toolchain and vcpkg. See README.md, "Build and test on Linux in Docker".
#
#   docker build -t ctrlexcel-build --build-arg BASE_IMAGE=winccoa321:3.21.7 docker
ARG BASE_IMAGE=winccoa321:3.21.7
FROM ${BASE_IMAGE}

USER root

RUN apt-get update \
 && apt-get install -y --no-install-recommends \
        build-essential \
        ca-certificates \
        cmake \
        curl \
        git \
        ninja-build \
        pkg-config \
        python3 \
        tar \
        unzip \
        zip \
 && apt-get clean && rm -rf /var/lib/apt/lists/*

# vcpkg at a fixed commit, so the image and its package cache are reproducible.
ARG VCPKG_COMMIT=07f4812200df3d3c931c0c8a6081d3b21fe2bf9f
RUN git clone https://github.com/microsoft/vcpkg.git /opt/vcpkg \
 && git -C /opt/vcpkg checkout "$VCPKG_COMMIT" \
 && /opt/vcpkg/bootstrap-vcpkg.sh -disableMetrics \
 && chown -R winccoa:winccoa /opt/vcpkg

COPY run-tests.sh /usr/local/bin/ctrlexcel-run-tests
RUN chmod a+rx /usr/local/bin/ctrlexcel-run-tests

# Mount points of the build directory and the vcpkg package cache: a new
# named volume takes over this ownership, so the winccoa user can write.
RUN mkdir -p /build /home/winccoa/.cache/vcpkg-archives \
 && chown -R winccoa:winccoa /build /home/winccoa/.cache

USER winccoa

ENV VCPKG_ROOT=/opt/vcpkg
ENV API_ROOT=/opt/WinCC_OA/3.21/api
ENV VCPKG_DEFAULT_BINARY_CACHE=/home/winccoa/.cache/vcpkg-archives

# The sources are mounted read-only at /src; build output goes to /build.
WORKDIR /build
ENTRYPOINT ["ctrlexcel-run-tests"]

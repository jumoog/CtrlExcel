# Build and test image for CtrlExcelReader on Linux: a WinCC OA image plus a
# C++ and a Rust toolchain. See README.md, "Build and test on Linux in Docker".
# The WinCC OA image is either Debian (Siemens' Dockerfile in the Debian
# package directory) or RHEL (docker/winccoa-rhel/Dockerfile):
#
#   docker build -t ctrlexcel-build --build-arg BASE_IMAGE=winccoa321:3.21.7 docker
#   docker build -t ctrlexcel-build-rhel --build-arg BASE_IMAGE=winccoa321-rhel:3.21.7 docker
ARG BASE_IMAGE=winccoa321:3.21.7
FROM ${BASE_IMAGE}

USER root

RUN if command -v apt-get > /dev/null; then \
      apt-get update \
      && apt-get install -y --no-install-recommends \
           build-essential ca-certificates cmake curl git ninja-build \
      && apt-get clean && rm -rf /var/lib/apt/lists/*; \
    else \
      dnf -y install --setopt=install_weak_deps=False \
           gcc-c++ make binutils ca-certificates cmake git ninja-build \
      && dnf clean all; \
    fi

COPY run-tests.sh /usr/local/bin/ctrlexcel-run-tests
RUN chmod a+rx /usr/local/bin/ctrlexcel-run-tests

# Mount points of the build directory, the cargo registry cache and the
# toolchains: a new named volume takes over this ownership, so the winccoa
# user can write.
RUN mkdir -p /build /home/winccoa/.cargo/registry /home/winccoa/.rustup \
 && chown -R winccoa:winccoa /build /home/winccoa/.cargo /home/winccoa/.rustup

USER winccoa

# rustup without a toolchain: run-tests.sh installs the one pinned in
# rust/rust-toolchain.toml into ~/.rustup (a volume), so a new pin (e.g.
# from Dependabot) needs no image rebuild and is downloaded once.
RUN curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs \
      | sh -s -- -y --no-modify-path --profile minimal --default-toolchain none
ENV PATH=/home/winccoa/.cargo/bin:$PATH

ENV API_ROOT=/opt/WinCC_OA/3.21/api

# The sources are mounted read-only at /src; build output goes to /build.
WORKDIR /build
ENTRYPOINT ["ctrlexcel-run-tests"]

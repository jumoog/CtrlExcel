# WinCC OA on RHEL 10 for building and testing CtrlExcelReader: what
# Siemens' Dockerfile for the Debian packages sets up (user winccoa,
# en_US.UTF-8), with only the packages a CTRL manager and the C++ API need.
# Siemens ships no Dockerfile with the RHEL packages. Build it with the
# package directory as context:
#
#   docker build -t winccoa321-rhel:3.21.7 -f docker/winccoa-rhel/Dockerfile path/to/WinCCOA-3.21.7-rhel.x86_64
#
# AlmaLinux is binary compatible with RHEL and, unlike UBI, has every
# library the packages need (with EPEL and CRB, e.g. ZeroMQ). A fixed minor
# release: its glibc is what the built .so requires.
ARG BASE_IMAGE=almalinux:10.3
FROM ${BASE_IMAGE}

ARG USER=winccoa
ARG USER_ID=1000
ARG USER_GID=1000

RUN dnf -y install epel-release \
 && dnf config-manager --set-enabled crb \
 && dnf -y install --setopt=install_weak_deps=False glibc-langpack-en sudo unzip \
 && dnf clean all

ENV LANG=en_US.UTF-8
ENV LANGUAGE=en_US:en
ENV LC_ALL=en_US.UTF-8

RUN groupadd --gid "$USER_GID" "$USER" \
 && useradd --uid "$USER_ID" --gid "$USER_GID" -m --home-dir "/home/$USER" "$USER" \
 && echo "$USER ALL=(ALL) NOPASSWD:ALL" >> /etc/sudoers

# CodeMeter provides libwibucm, Base the libraries, BaseServer WCCOActrl,
# APIcpp the API (api/CMakeDefines.txt).
COPY CodeMeter-*.rpm WinCCOA-*-Base-rhel.x86_64.rpm WinCCOA-*-BaseServer-rhel.x86_64.rpm WinCCOA-*-APIcpp-rhel.x86_64.rpm /tmp/rpm/
# Without weak dependencies: Base would pull in desktop and audio services.
RUN dnf -y install --setopt=install_weak_deps=False /tmp/rpm/*.rpm \
 && dnf clean all \
 && rm -rf /tmp/rpm

USER $USER
ENV OAINST=/opt/WinCC_OA/3.21

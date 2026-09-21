#!/bin/bash
# Build skeepto-engine once inside the manylinux container (CIBW_BEFORE_ALL).
# Static .a must be PIC so they can link into the Python .so.
set -euo pipefail

ENGINE="${SK_ENGINE_DIR:-/opt/skeepto-engine}"
mkdir -p "$(dirname "$ENGINE")"

if [[ ! -d "$ENGINE/.git" ]]; then
	if [[ -n "${SKEEPTO_ENGINE_TOKEN:-}" ]]; then
		git clone --depth 1 \
			"https://x-access-token:${SKEEPTO_ENGINE_TOKEN}@github.com/Stephane-76/SkeeptoEngine.git" \
			"$ENGINE"
	else
		git clone --depth 1 \
			"https://github.com/Stephane-76/SkeeptoEngine.git" \
			"$ENGINE"
	fi
fi

# AlmaLinux 8 (manylinux_2_28): C++20 toolchain + zip headers.
if command -v dnf >/dev/null 2>&1; then
	dnf install -y gcc-c++ git zlib-devel || true
	dnf install -y epel-release || true
	dnf install -y libzip-devel pugixml-devel || true
elif command -v yum >/dev/null 2>&1; then
	yum install -y gcc-c++ git zlib-devel || true
	yum install -y epel-release || true
	yum install -y libzip-devel pugixml-devel || true
fi

# EL8 system cmake is too old (3.11). Always drop a recent Kitware cmake.
if ! cmake --version 2>/dev/null | head -1 | grep -qE 'version 3\.(1[6-9]|[2-9][0-9])|version [4-9]'; then
	ver="3.28.6"
	curl -sSL "https://github.com/Kitware/CMake/releases/download/v${ver}/cmake-${ver}-linux-x86_64.tar.gz" \
		| tar -xz -C /usr/local --strip-components=1
fi

export CFLAGS="-fPIC"
export CXXFLAGS="-fPIC"

cmake -S "$ENGINE" -B "$ENGINE/build-unix" \
	-DSK_PLATFORM=unix \
	-DCMAKE_BUILD_TYPE=Release \
	-DSK_BUILD_TESTS=OFF \
	-DCMAKE_POSITION_INDEPENDENT_CODE=ON \
	-DCMAKE_C_FLAGS=-fPIC \
	-DCMAKE_CXX_FLAGS=-fPIC
cmake --build "$ENGINE/build-unix" --parallel

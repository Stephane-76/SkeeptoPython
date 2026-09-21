#!/bin/bash
# Build skeepto-engine once inside the manylinux container (CIBW_BEFORE_ALL).
# Static .a must be PIC so they can link into the Python .so.
set -eux

ENGINE="${SK_ENGINE_DIR:-/tmp/skeepto-engine}"
mkdir -p "$ENGINE"

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

# AlmaLinux 8 (manylinux_2_28): zip headers. C++ toolchain is already in the image.
if command -v yum >/dev/null 2>&1; then
	yum install -y zlib-devel || true
	yum install -y epel-release || true
	yum install -y libzip-devel pugixml-devel || true
fi

# Prefer the manylinux CPython pip (writable, no /usr/local tarball).
PYBIN=""
for wCand in /opt/python/cp312-cp312/bin /opt/python/cp313-cp313/bin; do
	if [[ -x "${wCand}/pip" ]]; then
		PYBIN="$wCand"
		break
	fi
done
if [[ -n "$PYBIN" ]]; then
	"${PYBIN}/pip" install cmake ninja
	export PATH="${PYBIN}:${PATH}"
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

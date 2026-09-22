#!/bin/bash
# Build skeepto-engine static libs inside the manylinux container (CIBW_BEFORE_ALL).
# Only the three libraries: the Python extension compiles SkExcel sources itself.
set -eux

# AlmaLinux 8 ships gcc 8, which cannot build this C++20 tree (make then exits 2).
if command -v yum >/dev/null 2>&1; then
	yum install -y epel-release
	yum install -y dnf-plugins-core || true
	yum config-manager --set-enabled powertools || true
	yum install -y git make cmake zlib-devel libzip-devel pugixml-devel \
		gcc-toolset-13 gcc-toolset-13-gcc-c++
fi

if [[ -f /opt/rh/gcc-toolset-13/enable ]]; then
	set +u
	# shellcheck disable=SC1091
	source /opt/rh/gcc-toolset-13/enable
	set -u
fi

echo "=== compiler $(command -v g++) ==="
g++ --version
g++ -dumpversion | awk -F. '{ if ($1 < 10) exit 1 }'

echo "=== cmake $(command -v cmake) ==="
cmake --version

ENGINE="${SK_ENGINE_DIR:-/tmp/skeepto-engine}"
if [[ ! -d "$ENGINE/.git" ]]; then
	rm -rf "$ENGINE"
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

export CFLAGS="-fPIC"
export CXXFLAGS="-fPIC"

cmake -S "$ENGINE" -B "$ENGINE/build-unix" \
	-DSK_PLATFORM=unix \
	-DCMAKE_BUILD_TYPE=Release \
	-DSK_BUILD_TESTS=OFF \
	-DSK_BUILD_APPS=OFF \
	-DCMAKE_POSITION_INDEPENDENT_CODE=ON \
	-DCMAKE_C_FLAGS=-fPIC \
	-DCMAKE_CXX_FLAGS=-fPIC
cmake --build "$ENGINE/build-unix" --parallel

test -f "$ENGINE/unix/lib/libSkSpreadSheet.a"
test -f "$ENGINE/unix/lib/libSkFormat.a"
test -f "$ENGINE/unix/lib/libSkRoot.a"

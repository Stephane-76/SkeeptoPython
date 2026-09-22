#!/bin/bash
# Build skeepto-engine static libs inside the manylinux container (CIBW_BEFORE_ALL).
# Only the three libraries: the Python extension compiles SkExcel sources itself.
set -eux

# Image PATH may omit the devtoolset when the entrypoint is not applied.
for wEnable in /opt/rh/gcc-toolset-14/enable /opt/rh/gcc-toolset-13/enable /opt/rh/gcc-toolset-12/enable; do
	if [[ -f "$wEnable" ]]; then
		# shellcheck disable=SC1090
		source "$wEnable"
		break
	fi
done

echo "=== compiler $(command -v g++ || true) ==="
g++ --version || true

# AlmaLinux 8 (manylinux_2_28): libzip / pugixml live in EPEL + powertools.
if command -v yum >/dev/null 2>&1; then
	yum install -y epel-release
	yum install -y dnf-plugins-core || true
	yum config-manager --set-enabled powertools || true
	yum install -y git gcc-c++ make cmake zlib-devel libzip-devel pugixml-devel
fi

echo "=== cmake $(command -v cmake || true) ==="
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

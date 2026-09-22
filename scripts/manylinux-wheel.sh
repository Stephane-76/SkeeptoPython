#!/bin/bash
# One-shot manylinux wheel. /io is skeepto-python, /engine is skeepto-engine.
# Run by docker so the compiler log is a normal Actions step (not buried in cibuildwheel).
set -eux

export SK_ENGINE_DIR="${SK_ENGINE_DIR:-/engine}"

bash /io/scripts/cibw-before-all.sh

# before-all installs and uses gcc-toolset-13 in its own shell. This shell
# still has the image default (gcc-toolset-14), which failed SkCore.cpp.
if [[ -f /opt/rh/gcc-toolset-13/enable ]]; then
	set +u
	# shellcheck disable=SC1091
	source /opt/rh/gcc-toolset-13/enable
	set -u
fi
export CC="$(command -v gcc)"
export CXX="$(command -v g++)"
echo "=== wheel compiler $($CXX) ==="
"$CXX" --version

PY=/opt/python/cp312-cp312/bin/python
"$PY" -m pip install -U pip wheel pytest auditwheel

export CMAKE_ARGS="-DSK_ENGINE_DIR=${SK_ENGINE_DIR} -DCMAKE_INTERPROCEDURAL_OPTIMIZATION=OFF"
export CFLAGS="-fPIC -fno-lto"
export CXXFLAGS="-fPIC -fno-lto"
export LDFLAGS="-fno-lto"

cd /io
# -v keeps the g++ diagnostics pip otherwise folds into "[N lines of output]".
"$PY" -m pip wheel -v . --no-deps --wheel-dir /tmp/wheels
mkdir -p /io/wheelhouse
auditwheel repair -w /io/wheelhouse /tmp/wheels/*.whl

"$PY" -m pip install /io/wheelhouse/*.whl
rm -rf /tmp/skeepto-tests
cp -a /io/tests /tmp/skeepto-tests
cd /tmp/skeepto-tests
"$PY" -m pytest -q

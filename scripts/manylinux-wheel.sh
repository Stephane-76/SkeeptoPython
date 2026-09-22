#!/bin/bash
# One-shot manylinux wheel. /io is skeepto-python, /engine is skeepto-engine.
# Run by docker so the compiler log is a normal Actions step (not buried in cibuildwheel).
set -eux

export SK_ENGINE_DIR="${SK_ENGINE_DIR:-/engine}"
bash /io/scripts/cibw-before-all.sh

PY=/opt/python/cp312-cp312/bin/python
"$PY" -m pip install -U pip wheel pytest auditwheel

export CMAKE_ARGS="-DSK_ENGINE_DIR=${SK_ENGINE_DIR} -DCMAKE_INTERPROCEDURAL_OPTIMIZATION=OFF"
export CFLAGS="-fPIC -fno-lto"
export CXXFLAGS="-fPIC -fno-lto"
export LDFLAGS="-fno-lto"

cd /io
"$PY" -m pip wheel . --no-deps --wheel-dir /tmp/wheels
mkdir -p /io/wheelhouse
auditwheel repair -w /io/wheelhouse /tmp/wheels/*.whl

"$PY" -m pip install /io/wheelhouse/*.whl
rm -rf /tmp/skeepto-tests
cp -a /io/tests /tmp/skeepto-tests
cd /tmp/skeepto-tests
"$PY" -m pytest -q

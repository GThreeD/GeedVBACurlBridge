# libcurl_vba_bridge

Small bridge for using libcurl from VBA on macOS.

## Build

```bash
cmake -S . -B build -G Ninja \
  -DCMAKE_TOOLCHAIN_FILE="$HOME/vcpkg/scripts/buildsystems/vcpkg.cmake"

cmake --build build
cmake --build build --target deploy
```

Build output is written to `out/`.

Import the files from `out/vba/` into your VBA project.

`--target deploy` copies libcurl_vba_bridge.dylib to:

```file
~/Library/Containers/com.microsoft.Excel/Data/curl/
```
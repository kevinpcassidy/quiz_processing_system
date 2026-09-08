# Poppler binary source and build provenance

Quiz Processing System (QPS) bundles a Windows Poppler distribution under
`vendor/poppler`. This document identifies the corresponding source and build
provenance. It does not change the MIT license of QPS.

## Distribution used by QPS

The bundled tree was taken from the version-specific
[`oschwartz10612/poppler-windows` release `v25.07.0-0`](https://github.com/oschwartz10612/poppler-windows/releases/tag/v25.07.0-0),
whose Windows archive is named `Release-25.07.0-0.zip`. The `-0` suffix is the
poppler-windows release/repackaging revision; it is not the conda-forge package
build number.

poppler-windows does not compile these Windows binaries independently. It
collects the Poppler package and its runtime dependencies from conda-forge and
repackages their files into a portable Windows ZIP. QPS in turn preserves the
relevant contents of that ZIP in `vendor/poppler`.

## Poppler package build

The Poppler files in this repository correspond to this exact conda-forge
artifact:

- Platform: `win-64`
- Package: `poppler`
- Version: `25.07.0`
- Build string: `h813ae87_1`
- Build number: `1`
- Full filename: [`poppler-25.07.0-h813ae87_1.conda`](https://anaconda.org/conda-forge/poppler/25.07.0/download/win-64/poppler-25.07.0-h813ae87_1.conda)

This identification is supported by the `25.07.0` version in the installed
headers and pkg-config files, the embedded conda build prefix
`D:/bld/poppler-split_1753302565761/_h_env`, and the matching July 23, 2025 PE
link timestamps. Thus, although the poppler-windows release is numbered
`25.07.0-0`, the Poppler conda package represented by the bundled binaries is
build **1**.

Representative SHA-256 checksums of the files currently bundled by QPS are:

```text
99e873ad35fd1a70ca8a6b85df53562d40088b6aa386fc37fb227db6a7d74739  vendor/poppler/Library/bin/poppler.dll
8ec2033f36247092866c757137d8815ab834d228e8322c45a820a12e14002f20  vendor/poppler/Library/bin/pdfinfo.exe
cdc7140e5f099312ee02d7b9ef13aa2c9924b829a57deb666ebf8295ca4b8af0  vendor/poppler/Library/lib/poppler.lib
```

## Corresponding source and recipe

The version-specific sources and build recipe are:

1. **poppler-windows packaging:** the
   [`v25.07.0-0` tag](https://github.com/oschwartz10612/poppler-windows/tree/v25.07.0-0)
   and its [release](https://github.com/oschwartz10612/poppler-windows/releases/tag/v25.07.0-0).
2. **Upstream Poppler source:**
   [`poppler-25.07.0.tar.xz`](https://poppler.freedesktop.org/poppler-25.07.0.tar.xz)
   from the [Poppler 25.07.0 release directory](https://poppler.freedesktop.org/releases.html).
3. **conda-forge recipe revision:** the recipe snapshot embedded in the exact
   [`win-64/poppler-25.07.0-h813ae87_1.conda`](https://anaconda.org/conda-forge/poppler/25.07.0/download/win-64/poppler-25.07.0-h813ae87_1.conda)
   artifact above. That immutable, build-specific artifact records the rendered
   recipe and source used by
   [`conda-forge/poppler-feedstock`](https://github.com/conda-forge/poppler-feedstock/tree/main/recipe).
4. **Poppler encoding data:** upstream
   [`poppler-data-0.4.12.tar.gz`](https://poppler.freedesktop.org/poppler-data-0.4.12.tar.gz).

The downloaded ZIP was unpacked before it was committed, and the original
conda `info/` records and `.conda` archives are not present in this repository.
For that reason, the exact conda artifact above is the permanent build-specific
reference for the rendered feedstock recipe. The installed `poppler-data.pc`
identifies the bundled data as version `0.4.12`; the unpacked tree does not
retain a conda build string for that data package.

## Licensing and preservation

The conda-forge package metadata designates Poppler as
`GPL-2.0-or-later`. QPS remains independently licensed under the MIT License;
including and invoking the separate Poppler utilities does not replace QPS's
license notice.

All upstream `LICENSE`, `COPYING`, `NOTICE`, README, authorship, and package
metadata files that are present under `vendor/poppler` and
`vendor_docs/poppler` must be retained when the application is redistributed.
In particular, the Poppler copying files and Xpdf notice are preserved under
`vendor/poppler/docs` and `vendor_docs/poppler/docs`, while the poppler-data
copying files are preserved under `vendor/poppler/share/poppler`.

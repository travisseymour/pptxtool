# SPDX-FileCopyrightText: PPTXTool | Copyright (c) 2026 Travis L. Seymour, PhD <nogard@ucsc.edu>
# SPDX-License-Identifier: GPL-3.0-or-later

from importlib.metadata import PackageNotFoundError, version

try:
    __version__ = version("pptxtool")
except PackageNotFoundError:  # running from a source tree that was never installed
    __version__ = "unknown"

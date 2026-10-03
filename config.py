"""
config.py – Centralized configuration loader for Google Maps Review Scraper.
Reads config.json and exposes a typed Config dataclass.
"""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
from pathlib import Path

logger = logging.getLogger(__name__)

ROOT_DIR = Path(__file__).parent
CONFIG_FILE = ROOT_DIR / "config.json"
OUTPUT_DIR = ROOT_DIR / "output"
LOGS_DIR = ROOT_DIR / "logs"
CHECKPOINTS_DIR = ROOT_DIR / "checkpoints"


@dataclass
class Config:
    google_maps_url: str = ""
    output_filename: str = "google_maps_reviews.xlsx"
    max_reviews: int = 0           # 0 = unlimited
    headless: bool = True
    scroll_pause: float = 2.5
    max_scroll_attempts_without_new_reviews: int = 10

    # Derived at runtime
    output_dir: Path = field(default_factory=lambda: OUTPUT_DIR)
    logs_dir: Path = field(default_factory=lambda: LOGS_DIR)
    checkpoints_dir: Path = field(default_factory=lambda: CHECKPOINTS_DIR)

    # Viewport / browser settings
    viewport_width: int = 1280
    viewport_height: int = 900
    locale: str = "id-ID"
    timezone: str = "Asia/Jakarta"
    user_agent: Optional[str] = None

    # Timeouts (ms)
    navigation_timeout: int = 60_000
    action_timeout: int = 15_000
    captcha_wait_timeout: int = 300_000   # 5 min

    # Checkpoint
    checkpoint_batch_size: int = 20

    def ensure_dirs(self) -> None:
        """Create output directories if they don't exist."""
        for d in (self.output_dir, self.logs_dir):
            d.mkdir(parents=True, exist_ok=True)

    @property
    def output_xlsx(self) -> Path:
        return self.output_dir / self.output_filename

    @property
    def output_csv(self) -> Path:
        stem = Path(self.output_filename).stem
        return self.output_dir / f"{stem}.csv"

    @property
    def checkpoint_file(self) -> Path:
        return self.checkpoints_dir / "reviews_checkpoint.json"


def load_config(path: Path = CONFIG_FILE) -> Config:
    """Load configuration from JSON file, fallback to defaults."""
    cfg = Config()

    if path.exists():
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
            for key, value in data.items():
                if hasattr(cfg, key):
                    setattr(cfg, key, value)
            logger.debug("Config loaded from %s", path)
        except (json.JSONDecodeError, OSError) as exc:
            logger.warning("Failed to load config.json: %s – using defaults.", exc)
    else:
        logger.info("config.json not found – using default configuration.")

    cfg.ensure_dirs()
    return cfg

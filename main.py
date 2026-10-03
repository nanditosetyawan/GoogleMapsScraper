"""
main.py – Entry point for Google Maps Review Scraper.

Usage:
    python main.py

Reads configuration from config.json automatically.
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path

from config import load_config, LOGS_DIR
from scraper import GoogleMapsReviewScraper


def setup_logging(logs_dir: Path) -> None:
    """Configure logging to both console and file."""
    logs_dir.mkdir(parents=True, exist_ok=True)
    log_file = logs_dir / "scraper.log"

    fmt = "[%(levelname)s] %(message)s"
    date_fmt = "%Y-%m-%d %H:%M:%S"

    handlers: list[logging.Handler] = [
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(log_file, encoding="utf-8"),
    ]

    logging.basicConfig(
        level=logging.INFO,
        format=fmt,
        datefmt=date_fmt,
        handlers=handlers,
        force=True,
    )

    # Reduce noise from playwright internal logs
    logging.getLogger("playwright").setLevel(logging.WARNING)
    logging.getLogger("asyncio").setLevel(logging.WARNING)


def main() -> int:
    # Load config first (creates dirs)
    cfg = load_config()

    # Setup logging
    setup_logging(cfg.logs_dir)

    logger = logging.getLogger(__name__)
    logger.info("=" * 50)
    logger.info("Google Maps Review Scraper – starting")
    logger.info("=" * 50)

    # Show config summary
    logger.info("URL        : %s", cfg.google_maps_url or "(akan ditanyakan)")
    logger.info("Output     : %s", cfg.output_filename)
    logger.info("Max reviews: %s", "Tanpa batas" if cfg.max_reviews == 0 else str(cfg.max_reviews))
    logger.info("Headless   : %s", cfg.headless)
    logger.info("Scroll pause: %.1fs", cfg.scroll_pause)
    logger.info("-" * 50)

    try:
        scraper = GoogleMapsReviewScraper(cfg)
        reviews = scraper.run()
        logger.info("Selesai. %d review dikumpulkan.", len(reviews))
        return 0
    except KeyboardInterrupt:
        logger.info("Scraping dibatalkan oleh user (Ctrl+C).")
        return 1
    except ValueError as e:
        logger.error("Kesalahan konfigurasi: %s", e)
        logger.error("Troubleshooting:")
        logger.error("  1. Pastikan URL di config.json valid.")
        logger.error("  2. Format: https://maps.app.goo.gl/... atau https://www.google.com/maps/...")
        return 2
    except RuntimeError as e:
        logger.error("Scraping gagal: %s", e)
        logger.error("Troubleshooting:")
        logger.error("  1. Cek file screenshot/HTML di folder logs/")
        logger.error("  2. Pastikan koneksi internet aktif.")
        logger.error("  3. Coba jalankan ulang – Google Maps mungkin mengubah struktur halaman.")
        logger.error("  4. Jika CAPTCHA: selesaikan di browser lalu jalankan ulang.")
        return 3
    except Exception as e:
        logger.error("Error tidak terduga: %s", e, exc_info=True)
        logger.error("Cek logs/scraper.log untuk detail lengkap.")
        return 4


if __name__ == "__main__":
    sys.exit(main())

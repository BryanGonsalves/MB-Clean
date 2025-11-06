"""Optional OpenAI-powered helpers for name and date normalization."""

from __future__ import annotations

import json
import os
from typing import Iterable, List, Optional


class AIUnavailableError(RuntimeError):
    """Raised when the OpenAI client cannot be initialised."""


class AINormalizer:
    """Thin wrapper around the OpenAI API for optional cleanup assistance."""

    def __init__(self, client, model: str = "gpt-4o-mini"):
        self._client = client
        self._model = model

    @classmethod
    def from_api_key(cls, api_key: Optional[str]) -> "AINormalizer":
        """Create a normalizer with an explicit API key."""

        if not api_key:
            raise AIUnavailableError("OpenAI API key is required to enable AI normalization.")

        try:
            from openai import OpenAI
        except ImportError as exc:
            raise AIUnavailableError("The openai package is not installed.") from exc

        client = OpenAI(api_key=api_key)
        return cls(client)

    @classmethod
    def from_env(cls) -> "AINormalizer":
        """Create a normalizer by reading OPENAI_API_KEY from the environment."""

        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            raise AIUnavailableError("OPENAI_API_KEY is not set.")
        return cls.from_api_key(api_key)

    def normalize_names(self, values: Iterable[str]) -> List[str]:
        """Return title-cased names using the OpenAI API when available."""

        values_list = list(values)
        if not values_list:
            return []

        payload = json.dumps(values_list)
        try:
            response = self._client.chat.completions.create(
                model=self._model,
                temperature=0,
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "You clean data values for an internal tool. "
                            "Return a JSON array of strings. "
                            "Fix capitalization, spacing, and punctuation for personal names."
                        ),
                    },
                    {
                        "role": "user",
                        "content": f"Normalize these names: {payload}",
                    },
                ],
            )
            content = response.choices[0].message.content.strip()
            cleaned = json.loads(content)
            if isinstance(cleaned, list) and len(cleaned) == len(values_list):
                return [str(item).strip() for item in cleaned]
        except Exception:
            pass

        return values_list

    def normalize_dates(self, values: Iterable[str]) -> List[str]:
        """Attempt to coerce dates into ISO 8601 strings via OpenAI."""

        values_list = list(values)
        if not values_list:
            return []

        payload = json.dumps(values_list)
        try:
            response = self._client.chat.completions.create(
                model=self._model,
                temperature=0,
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "You convert messy date strings into the ISO 8601 format YYYY-MM-DD. "
                            "Return a JSON array of strings. "
                            "If a value cannot be converted confidently, echo it unchanged."
                        ),
                    },
                    {
                        "role": "user",
                        "content": f"Normalize these dates: {payload}",
                    },
                ],
            )
            content = response.choices[0].message.content.strip()
            cleaned = json.loads(content)
            if isinstance(cleaned, list) and len(cleaned) == len(values_list):
                return [str(item).strip() for item in cleaned]
        except Exception:
            pass

        return values_list

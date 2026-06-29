from functools import lru_cache
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    telegram_token: str
    admin_ids: str = ""
    bot_db_path: str = "/data/uzum_bot.db"
    uzum_token: str = ""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    @property
    def allowed_user_ids(self) -> set[int]:
        ids: set[int] = set()
        for raw_id in self.admin_ids.split(","):
            raw_id = raw_id.strip()
            if raw_id.isdigit():
                ids.add(int(raw_id))
        return ids


@lru_cache
def get_settings() -> Settings:
    return Settings()


settings = get_settings()

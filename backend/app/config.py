from pydantic_settings import BaseSettings, SettingsConfigDict

class Settings(BaseSettings):
    app_name: str = "CRM Operacional AmPm"
    environment: str = "production"
    mongo_uri: str = "mongodb://localhost:27017"
    mongo_db: str = "crm_ampm"
    jwt_secret: str = "change-me-in-production"
    jwt_expire_minutes: int = 480
    frontend_origin: str = "*"

    model_config = SettingsConfigDict(env_file=".env", extra="ignore")

settings = Settings()

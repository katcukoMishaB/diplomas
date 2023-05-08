from pydantic import BaseSettings


class Settings(BaseSettings):
    mail_server: str
    mail_port: int
    mail_login: str
    mail_password: str


settings = Settings(_env_file=".env", _env_file_encoding="utf-8")
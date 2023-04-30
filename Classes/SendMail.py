import aiosmtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import asyncio


from typing import List


class SendMail:
    def __init__(self,
                 server: str,
                 post: int,
                 login: str,
                 password: str):
        self.__server: str = server
        self.__post: int = post
        self.__login: str = login
        self.__password: str = password

    async def __send_message(self, message: MIMEMultipart):
        async with aiosmtplib.SMTP(self.__server, self.__post) as smtp:
            await smtp.login(self.__login, self.__password)
            await smtp.send_message(message)

    def send_message(self, list_message: List[MIMEMultipart]):
        loop = asyncio.get_event_loop()
        funct = [self.__send_message(message) for message in list_message]
        loop.run_until_complete(asyncio.gather(*funct))

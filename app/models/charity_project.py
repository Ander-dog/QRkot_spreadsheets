from sqlalchemy import Column, String, Text

from app.models.base import Receipt


class CharityProject(Receipt):
    name = Column(String(100), unique=True, nullable=False)
    description = Column(Text, nullable=False)

    def __repr__(self):
        if self.fully_invested:
            return (f'Сбор для {self.name} закрыт. Было собранно '
                    f'{self.full_amount}.')
        return (f'Для {self.name} уже собранно {self.invested_amount} из '
                f'{self.full_amount}.')

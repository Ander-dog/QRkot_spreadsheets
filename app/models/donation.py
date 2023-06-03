from sqlalchemy import Column, ForeignKey, Integer, Text

from app.models.base import Receipt


class Donation(Receipt):
    user_id = Column(Integer, ForeignKey('user.id'), nullable=False)
    comment = Column(Text)

    def __repr__(self):
        if self.fully_invested:
            return (f'Вся сумма пожертвования {self.full_amount} была '
                    'инвестированна в проекты')
        return (f'Из {self.full_amount} этого пожертвования было '
                f'инвестированно {self.invested_amount}.')

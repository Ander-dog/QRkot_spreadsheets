from datetime import datetime
from typing import Optional

from pydantic import (BaseModel, Extra, Field, NonNegativeInt, PositiveInt,
                      validator)


class CharityProjectBase(BaseModel):
    name: Optional[str] = Field(None, min_length=1, max_length=100)
    description: Optional[str] = Field(None, min_length=1)
    full_amount: Optional[PositiveInt]

    class Config:
        extra = Extra.forbid


class CharityProjectCreate(CharityProjectBase):
    name: str = Field(..., min_length=1, max_length=100)
    description: str = Field(..., min_length=1)
    full_amount: PositiveInt


class CharityProjectUpdate(CharityProjectBase):

    @validator('name')
    def name_cant_be_null(cls, value):
        if value is None:
            raise ValueError('У проекта обязательно должно быть имя')
        return value

    @validator('description')
    def description_cant_be_null(cls, value):
        if value is None:
            raise ValueError('У проекта обязательно должно быть описание')
        return value

    @validator('full_amount')
    def full_amount_cant_be_null(cls, value):
        if value is None:
            raise ValueError('У проекта обязательно должна быть конечная цель')
        return value


class CharityProjectDB(CharityProjectBase):
    id: int
    invested_amount: NonNegativeInt
    fully_invested: bool
    create_date: datetime
    close_date: Optional[datetime]

    class Config:
        orm_mode = True
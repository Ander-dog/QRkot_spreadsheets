from datetime import datetime
from typing import Union

from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from app.models import CharityProject, Donation

MODEL_PAIRS = {
    Donation: CharityProject,
    CharityProject: Donation,
}


def close(
        closing_obj: Union[CharityProject, Donation],
        session: AsyncSession
) -> None:
    closing_obj.invested_amount = closing_obj.full_amount
    closing_obj.fully_invested = True
    closing_obj.close_date = datetime.now()
    session.add(closing_obj)


def investment(
        invest_obj: Union[CharityProject, Donation],
        pair_obj: Union[Donation, CharityProject],
        session: AsyncSession
) -> None:
    invest_money = invest_obj.full_amount - invest_obj.invested_amount
    pair_money = pair_obj.full_amount - pair_obj.invested_amount
    if invest_money > pair_money:
        invest_obj.invested_amount += pair_money
        session.add(invest_obj)
        close(pair_obj, session)
    elif invest_money < pair_money:
        pair_obj.invested_amount += invest_money
        session.add(pair_obj)
        close(invest_obj, session)
    else:
        close(invest_obj, session)
        close(pair_obj, session)


async def invest_in_model(
        invest_obj: Union[CharityProject, Donation],
        session: AsyncSession
) -> None:
    pair_model = MODEL_PAIRS[type(invest_obj)]
    while not invest_obj.fully_invested:
        pair_obj = await session.execute(
            select(pair_model).where(pair_model.fully_invested.is_(False))
        )
        pair_obj = pair_obj.scalars().first()
        if pair_obj is None:
            break
        investment(invest_obj, pair_obj, session)

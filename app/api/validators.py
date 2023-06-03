from http import HTTPStatus
from typing import Optional

from fastapi import HTTPException
from sqlalchemy.ext.asyncio import AsyncSession

from app.crud import charityproject_crud
from app.models import CharityProject


async def check_name_duplicate(
        project_name: str,
        session: AsyncSession,
) -> None:
    project = await charityproject_crud.get_by_attribute(
        'name',
        project_name,
        session,
    )
    if project:
        raise HTTPException(
            status_code=HTTPStatus.BAD_REQUEST,
            detail='Проект с таким именем уже существует!',
        )


async def check_charityproject_exists(
        charityproject_id: int,
        session: AsyncSession,
) -> CharityProject:
    charityproject = await charityproject_crud.get(charityproject_id, session)
    if charityproject is None:
        raise HTTPException(
            status_code=HTTPStatus.NOT_FOUND,
            detail='Проект не найден'
        )
    return charityproject


async def check_charityproject_for_delete(
        charityproject: CharityProject,
) -> CharityProject:
    if charityproject.invested_amount != 0:
        raise HTTPException(
            status_code=HTTPStatus.BAD_REQUEST,
            detail=('В проект были внесены средства, не подлежит удалению!')
        )
    return charityproject


async def check_charityproject_for_update(
        charityproject: CharityProject,
        new_goal: Optional[int] = None,
) -> CharityProject:
    if charityproject.fully_invested:
        raise HTTPException(
            status_code=HTTPStatus.BAD_REQUEST,
            detail=('Закрытый проект нельзя редактировать!')
        )
    if new_goal is not None and charityproject.invested_amount > new_goal:
        raise HTTPException(
            status_code=HTTPStatus.BAD_REQUEST,
            detail=('Нельзя ставить цель меньше уже внесённых средств')
        )
    return charityproject

from typing import List

from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.api.validators import (check_charityproject_exists,
                                check_charityproject_for_delete,
                                check_charityproject_for_update,
                                check_name_duplicate)
from app.core.db import get_async_session
from app.core.user import current_superuser
from app.crud import charityproject_crud
from app.models import CharityProject
from app.schemas.charityproject import (CharityProjectCreate, CharityProjectDB,
                                        CharityProjectUpdate)

router = APIRouter()


@router.post(
    '/',
    response_model=CharityProjectDB,
    dependencies=[Depends(current_superuser)],
    response_model_exclude_none=True,
)
async def create_new_charityproject(
        charityproject: CharityProjectCreate,
        session: AsyncSession = Depends(get_async_session),
) -> CharityProject:
    """Только для суперюзеров."""
    await check_name_duplicate(charityproject.name, session)
    new_project = await charityproject_crud.create(charityproject, session)
    return new_project


@router.get(
    '/',
    response_model=List[CharityProjectDB],
    response_model_exclude_none=True,
)
async def get_all_charityprojects(
        session: AsyncSession = Depends(get_async_session),
) -> List[CharityProject]:
    all_projects = await charityproject_crud.get_multi(session)
    return all_projects


@router.patch(
    '/{charityproject_id}',
    response_model=CharityProjectDB,
    dependencies=[Depends(current_superuser)],
)
async def partially_update_charityproject(
        charityproject_id: int,
        obj_in: CharityProjectUpdate,
        session: AsyncSession = Depends(get_async_session),
) -> CharityProject:
    """Только для суперюзеров."""
    charityproject = await check_charityproject_exists(
        charityproject_id,
        session,
    )

    if obj_in.name is not None:
        await check_name_duplicate(obj_in.name, session)

    await check_charityproject_for_update(
        charityproject,
        obj_in.full_amount,
    )

    charityproject = await charityproject_crud.update(
        charityproject, obj_in, session
    )
    return charityproject


@router.delete(
    '/{charityproject_id}',
    response_model=CharityProjectDB,
    dependencies=[Depends(current_superuser)],
)
async def remove_charityproject(
        charityproject_id: int,
        session: AsyncSession = Depends(get_async_session),
) -> CharityProject:
    """Только для суперюзеров."""
    charityproject = await check_charityproject_exists(
        charityproject_id,
        session,
    )
    charityproject = await check_charityproject_for_delete(charityproject)
    charityproject = await charityproject_crud.remove(
        charityproject, session
    )
    return charityproject
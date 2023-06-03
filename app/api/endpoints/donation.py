from typing import List

from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.db import get_async_session
from app.core.user import current_superuser, current_user
from app.crud import donation_crud
from app.models import Donation, User
from app.schemas.donation import DonationCreate, DonationDB, DonationFullDB

router = APIRouter()


@router.post(
    '/',
    response_model=DonationDB,
    response_model_exclude_none=True,
)
async def create_donation(
        donation: DonationCreate,
        session: AsyncSession = Depends(get_async_session),
        user: User = Depends(current_user),
) -> Donation:
    new_donation = await donation_crud.create(donation, session, user)
    return new_donation


@router.get(
    '/',
    response_model=List[DonationFullDB],
    dependencies=[Depends(current_superuser)],
    response_model_exclude_none=True,
)
async def get_all_donations(
        session: AsyncSession = Depends(get_async_session),
) -> List[Donation]:
    """Только для суперюзеров."""
    all_donations = await donation_crud.get_multi(session)
    return all_donations


@router.get(
    '/my',
    response_model=List[DonationDB],
    response_model_exclude_none=True,
)
async def get_my_donations(
        session: AsyncSession = Depends(get_async_session),
        user: User = Depends(current_user),
) -> List[Donation]:
    user_donations = await donation_crud.get_by_attribute(
        attr_name='user_id',
        attr_value=user.id,
        session=session
    )
    return user_donations

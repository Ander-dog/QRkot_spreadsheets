from typing import List

from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy.sql import func

from app.crud.base import CRUDBase
from app.models.charity_project import CharityProject
from app.schemas.charityproject import (CharityProjectCreate,
                                        CharityProjectUpdate)


class CRUDCharityProject(CRUDBase[
    CharityProject,
    CharityProjectCreate,
    CharityProjectUpdate
]):

    async def get_projects_by_completion_rate(
            self,
            session: AsyncSession,
    ) -> List[CharityProject]:
        projects = await session.execute(
            select(CharityProject).where(
                CharityProject.fully_invested.is_(True)
            ).order_by(
                (func.julianday(CharityProject.close_date) -
                 func.julianday(CharityProject.create_date))
            )
        )
        projects = projects.scalars().all()
        return projects


charityproject_crud = CRUDCharityProject(CharityProject)

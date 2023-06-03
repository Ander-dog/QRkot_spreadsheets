from app.crud.base import CRUDBase
from app.models.donation import Donation
from app.schemas.donation import DonationCreate, DonationUpdate


class CRUDdonation(CRUDBase[
    Donation,
    DonationCreate,
    DonationUpdate
]):
    pass


donation_crud = CRUDdonation(Donation)

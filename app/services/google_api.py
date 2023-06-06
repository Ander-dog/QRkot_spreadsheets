from datetime import datetime
from typing import List, Dict, Union, Any

from aiogoogle import Aiogoogle

from app.core.config import settings
from app.models import CharityProject

FORMAT = "%Y/%m/%d %H:%M:%S"


def generate_spreadsheet_config(
    time: datetime = None
) -> Dict[str, Union[Dict[str, str], List[Dict]]]:
    if time is None:
        time = datetime.now().strftime(FORMAT)
    spreadsheet_body = {
        'properties': {'title': f'Отчет по инвестициям на {time}',
                       'locale': 'ru_RU'},
        'sheets': [{'properties': {'sheetType': 'GRID',
                                   'sheetId': 0,
                                   'title': 'Лист1',
                                   'gridProperties': {'rowCount': 100,
                                                      'columnCount': 3}}}]
    }
    return spreadsheet_body


def generate_table_header(
    time: datetime = None
) -> List[List[Union[str, datetime]]]:
    if time is None:
        time = datetime.now().strftime(FORMAT)
    header_table_values = [
        ['Отчет от', time],
        ['Топ проектов по скорости закрытия'],
        ['Название проекта', 'Время сбора', 'Описание'],
    ]
    return header_table_values


def generate_permissions_body(email: str) -> Dict[str, str]:
    permissions_body = {
        'type': 'user',
        'role': 'writer',
        'emailAddress': email
    }
    return permissions_body


def generate_request_body(
    values: List[List[Any]]
) -> Dict[str, Union[str, List[List[Any]]]]:
    request_body = {
        'majorDimension': 'ROWS',
        'values': values
    }
    return request_body


async def spreadsheets_create(wrapper_services: Aiogoogle) -> str:
    service = await wrapper_services.discover('sheets', 'v4')
    spreadsheet_body = generate_spreadsheet_config()
    response = await wrapper_services.as_service_account(
        service.spreadsheets.create(json=spreadsheet_body)
    )
    spreadsheet_id = response['spreadsheetId']
    return spreadsheet_id


async def set_user_permissions(
        spreadsheet_id: str,
        wrapper_services: Aiogoogle
) -> None:
    permissions_body = generate_permissions_body(settings.email)
    service = await wrapper_services.discover('drive', 'v3')
    await wrapper_services.as_service_account(
        service.permissions.create(
            fileId=spreadsheet_id,
            json=permissions_body,
            fields='id'
        ))


async def spreadsheets_update_value(
        spreadsheet_id: str,
        charity_projects: List[CharityProject],
        wrapper_services: Aiogoogle
) -> None:
    service = await wrapper_services.discover('sheets', 'v4')
    table_values = generate_table_header()
    for project in charity_projects:
        new_row = [
            project.name,
            str(project.close_date - project.create_date),
            project.description
        ]
        table_values.append(new_row)

    append_body = generate_request_body(table_values)
    sheet_range = 'A1:C' + str(len(table_values))
    await wrapper_services.as_service_account(
        service.spreadsheets.values.append(
            spreadsheetId=spreadsheet_id,
            range=sheet_range,
            valueInputOption='USER_ENTERED',
            insertDataOption='OVERWRITE',
            json=append_body
        )
    )

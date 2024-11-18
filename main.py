from typing import List, Tuple, Optional
import tableauserverclient as TSC
from dotenv import load_dotenv
import pandas as pd
import numpy as np
import argparse
import os


load_dotenv()


class Connection:
    def __init__(self, server_url: str, username: str, password: str, site_id: str = ''):
        self.server = self.connect_to_tableau(
            server_url, username, password, site_id)

    def connect_to_tableau(self, server_url: str, username: str, password: str, site_id: str = ''):
        tableau_auth = TSC.TableauAuth(username, password, site_id)
        server = TSC.Server(server_url, use_server_version=True)
        try:
            print("Attempting to sign in...")
            server.auth.sign_in(tableau_auth)
            print("Signed in successfully!")
            return server
        except Exception as e:
            print(f"Failed to sign in: {e}")
            raise

    def __del__(self):
        if self.server and self.server.is_signed_in():
            try:
                print("Signing out...")
                self.server.auth.sign_out()
                print("Signed out successfully!")
            except Exception as e:
                print(f"Failed to sign out: {e}")
                raise


class Permission(Connection):
    def __init__(self, server_url, username, password, site_id=''):
        super().__init__(server_url, username, password, site_id)

    def list_users(self) -> pd.DataFrame:
        try:
            users, _ = self.server.users.get()
            user_list = []
            for user in users:
                if user.site_role != 'Unlicensed':
                    user_list.append(user.name)
            df = pd.DataFrame({'user_name': user_list})
            df['Name'] = 'All Users'
            return df
        except Exception as e:
            print(f"Failed to retrieve users: {e}")
            raise

    def get_permissions_and_metadata_for_item(self, item: str, item_type: str) -> Tuple[
        List[str], List[str], List[str], Optional[str], Optional[str]
    ]:
        try:
            user_types, user_names, user_capabilities = [], [], []
            getattr(self.server, item_type).populate_permissions(item)
            create_date = getattr(item, 'created_at', None)
            create_date = create_date.strftime(
                '%Y-%m-%d %H:%M:%S') if create_date else None
            owner = self.server.users.get_by_id(item.owner_id)
            owner = owner.fullname if owner else None
            permissions = item.permissions

            for rule in permissions:
                group_user_type = rule.grantee.tag_name
                group_user_id = rule.grantee.id
                group_user_capabilities = rule.capabilities.get(
                    'Read', 'No permissions')
                if group_user_type == 'user':
                    user_item = self.server.users.get_by_id(group_user_id)
                    group_user_name = user_item.name
                elif group_user_type == 'group':
                    for group_item in TSC.Pager(self.server.groups):
                        if group_item.id == group_user_id:
                            group_user_name = group_item.name
                            break
                user_types.append(group_user_type)
                user_names.append(group_user_name)
                user_capabilities.append(group_user_capabilities)
            return (user_types, user_names, user_capabilities, create_date, owner)

        except Exception as e:
            print(f"An error occurred: {e}")
            user_names = ["Deny for all users"]
            user_capabilities = ['Allow']
            return [['Deny'], user_names, user_capabilities] + ['Deny']*2

    def fetch_and_list_permissions(self) -> pd.DataFrame:
        project = [p for p in TSC.Pager(self.server.projects)]
        sources = {'projects', 'workbooks', 'datasources', 'flows'}
        projects = []

        for attr in sources:
            if attr == 'projects':
                for p in project:
                    user_type, user_name, user_capabilities, create_date, owner = self.get_permissions_and_metadata_for_item(
                        p, attr)
                    for group_user_type, group_user_name, group_user_capabilities in zip(user_type, user_name, user_capabilities):
                        pj_data = {
                            "Object Type": attr,
                            "Project/Folder Name/Dashboard Name": p.name,
                            "Type": group_user_type,
                            "Name": group_user_name,
                            "Read_Permission": group_user_capabilities,
                            "Create_Date": create_date,
                            "Owner": owner
                        }
                        projects.append(pj_data)
            else:
                all_datas, pagination_item = getattr(self.server, attr).get()
                project_datas = [data for data in all_datas]
                for data_type in project_datas:
                    user_type, user_name, user_capabilities, create_date, owner = self.get_permissions_and_metadata_for_item(
                        data_type, attr)
                    for group_user_type, group_user_name, group_user_capabilities in zip(user_type, user_name, user_capabilities):
                        pj_data = {
                            "Object Type": attr,
                            "Project/Folder Name/Dashboard Name": data_type.name,
                            "Type": group_user_type,
                            "Name": group_user_name,
                            "Read_Permission": group_user_capabilities,
                            "Create_Date": create_date,
                            "Owner": owner
                        }
                        projects.append(pj_data)
        return pd.DataFrame(projects)


def main():
    # Tableau Server credentials
    server_url = os.getenv('TABLEAU_SERVER_URL')
    username = os.getenv('TABLEAU_USERNAME')
    password = os.getenv('TABLEAU_PASSWORD')
    site_id = os.getenv('TABLEAU_SITE_ID', '')

    parser = argparse.ArgumentParser(
        description='View permissions for all items in a project')
    parser.add_argument('--siteId', '-s', required=False,
                        help='project area not including stage')
    parser.add_argument('--output', '-o', required=False, default='output.xlsx',
                        help='output file name for Excel')
    args = parser.parse_args()
    site_id = args.siteId if args.siteId else ''
    permission = Permission(server_url, username, password, site_id)
    df = permission.fetch_and_list_permissions()
    user_df = permission.list_users()
    df_merged = pd.merge(df, user_df, on='Name', how='left')
    df_merged['user_name'] = df_merged['user_name'].fillna(df_merged['Name'])
    df = df_merged.drop(columns=['Type', 'Name'])
    col = list(df.columns.difference(['Read_Permission']))
    df = df.sort_values(by=col + ['Read_Permission'],
                        ascending=[True]*(len(col)+1))
    df = df.drop_duplicates(subset=col, keep='first')
    df['Create_Date'] = df['Create_Date'].fillna('Deny')
    df = df.pivot_table(index=['Object Type', 'Project/Folder Name/Dashboard Name',
                               'Create_Date', 'Owner'], columns='user_name', values='Read_Permission', aggfunc=lambda x: ','.join(x)).reset_index()
    df['No.'] = df.groupby('Object Type').cumcount() + 1
    df['Site'] = args.siteId if args.siteId else 'Default'
    df['Objective/Purpose'] = np.nan
    columns = ['Site', 'No.', 'Project/Folder Name/Dashboard Name',
               'Objective/Purpose', 'Object Type', 'Owner', 'Create_Date']
    col = columns + list(df.columns.difference(columns))
    df = df[col]
    df = df.replace({'Allow': '✓', 'Deny': np.nan})
    df['Create_Date'] = pd.to_datetime(df['Create_Date']).dt.date
    df = df.reset_index(drop=True)
    with pd.ExcelWriter(args.output, mode='w', engine='xlsxwriter') as writer:
        sheet_name = "Permissions"
        df.to_excel(writer, sheet_name=sheet_name,
                    startrow=1, header=False, index=False)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': False,
            'valign': 'center',
            'fg_color': '#ADD8E6',
            'font_size': 12,
            'border': 1})

        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        worksheet.autofit()


if __name__ == '__main__':
    main()

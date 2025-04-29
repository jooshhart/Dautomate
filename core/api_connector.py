# from office365.sharepoint.client_context import ClientContext
# from office365.runtime.auth.authentication_context import AuthenticationContext

# class APIConnector:
#     def __init__(self, site_url, username, password):
#         self.ctx = None
#         self.site_url = site_url
#         self.username = username
#         self.password = password

#     def connect(self):
#         ctx_auth = AuthenticationContext(self.site_url)
#         if ctx_auth.acquire_token_for_user(self.username, self.password):
#             self.ctx = ClientContext(self.site_url, ctx_auth)
#             print("Connected to SharePoint")
#         else:
#             print("Authentication failed")

#     def upload_file(self, local_path, target_folder):
#         with open(local_path, 'rb') as f:
#             target_folder = self.ctx.web.get_folder_by_server_relative_url(target_folder)
#             target_folder.upload_file(local_path.split("/")[-1], f.read())
#             self.ctx.execute_query()
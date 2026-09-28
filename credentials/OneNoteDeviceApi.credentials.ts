import type {
	IAuthenticateGeneric,
	ICredentialDataDecryptedObject,
	ICredentialTestRequest,
	ICredentialType,
	IDataObject,
	IHttpRequestHelper,
	Icon,
	INodeProperties,
} from 'n8n-workflow';
import { authority, FORM_HEADERS, formBody, scopes } from '../nodes/OneNote/utils/auth';

/**
 * Alternative to the OAuth2 credential for setups where the n8n redirect URL cannot be
 * registered in Azure (http on a LAN IP, no HTTPS). The refresh token is created once with
 * the "OneNote Login Helper" node (device code flow) and renewed automatically afterwards.
 */
export class OneNoteDeviceApi implements ICredentialType {
	name = 'oneNoteDeviceApi';

	displayName = 'OneNote Device Login API';

	documentationUrl =
		'https://github.com/petopp/n8n-onenote-system/blob/main/docs/setup-azure.md#ohne-redirect-uri-device-login';

	icon: Icon = {
		light: 'file:../nodes/OneNote/onenote.svg',
		dark: 'file:../nodes/OneNote/onenote.dark.svg',
	};

	properties: INodeProperties[] = [
		{
			displayName: 'Client ID',
			name: 'clientId',
			type: 'string',
			default: '',
			required: true,
			description:
				'Application (client) ID of your Azure app registration. "Allow public client flows" must be enabled.',
		},
		{
			displayName: 'Account Type',
			name: 'tenant',
			type: 'options',
			options: [
				{ name: 'Personal Microsoft Account', value: 'consumers' },
				{ name: 'Personal + Work/School Accounts', value: 'common' },
			],
			default: 'consumers',
		},
		{
			displayName: 'Enable Video Upload (OneDrive)',
			name: 'enableOneDrive',
			type: 'boolean',
			default: true,
			description: 'Must match the choice made in the OneNote Login Helper node',
		},
		{
			displayName: 'Refresh Token',
			name: 'refreshToken',
			type: 'string',
			typeOptions: { password: true },
			default: '',
			required: true,
			description: 'Created with the OneNote Login Helper node. Renewed automatically.',
		},
		{
			displayName: 'Access Token',
			name: 'accessToken',
			type: 'hidden',
			typeOptions: { expirable: true, password: true },
			default: '',
		},
	];

	async preAuthentication(
		this: IHttpRequestHelper,
		credentials: ICredentialDataDecryptedObject,
	): Promise<IDataObject> {
		const response = (await this.helpers.httpRequest({
			method: 'POST',
			url: `${authority(String(credentials.tenant))}/token`,
			headers: FORM_HEADERS,
			body: formBody({
				client_id: String(credentials.clientId),
				grant_type: 'refresh_token',
				refresh_token: String(credentials.refreshToken),
				scope: scopes(credentials.enableOneDrive !== false),
			}),
			json: true,
		})) as IDataObject;
		return {
			accessToken: response.access_token as string,
			// Microsoft rotates refresh tokens; keep the newest one so the login does not expire.
			...(response.refresh_token ? { refreshToken: response.refresh_token as string } : {}),
		};
	}

	authenticate: IAuthenticateGeneric = {
		type: 'generic',
		properties: { headers: { Authorization: '=Bearer {{$credentials.accessToken}}' } },
	};

	test: ICredentialTestRequest = {
		request: {
			baseURL: 'https://graph.microsoft.com/v1.0',
			url: '/me/onenote/notebooks',
			qs: { $top: 1 },
		},
	};
}

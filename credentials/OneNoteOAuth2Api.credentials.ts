import type { ICredentialType, Icon, INodeProperties } from 'n8n-workflow';

export class OneNoteOAuth2Api implements ICredentialType {
	name = 'oneNoteOAuth2Api';

	extends = ['oAuth2Api'];

	displayName = 'OneNote OAuth2 API';

	icon: Icon = {
		light: 'file:../nodes/OneNote/onenote.svg',
		dark: 'file:../nodes/OneNote/onenote.dark.svg',
	};

	documentationUrl = 'https://github.com/petopp/n8n-onenote-system/blob/main/docs/setup-azure.md';

	properties: INodeProperties[] = [
		{
			displayName: 'Grant Type',
			name: 'grantType',
			type: 'hidden',
			default: 'authorizationCode',
		},
		{
			displayName: 'Account Type',
			name: 'tenant',
			type: 'options',
			options: [
				{
					name: 'Personal Microsoft Account',
					value: 'consumers',
					description: 'Outlook.com, Hotmail, Live or Xbox account',
				},
				{
					name: 'Personal + Work/School Accounts',
					value: 'common',
					description: 'The app registration must allow both account types',
				},
			],
			default: 'consumers',
		},
		{
			displayName: 'Authorization URL',
			name: 'authUrl',
			type: 'hidden',
			default: '=https://login.microsoftonline.com/{{$self["tenant"]}}/oauth2/v2.0/authorize',
			required: true,
		},
		{
			displayName: 'Access Token URL',
			name: 'accessTokenUrl',
			type: 'hidden',
			default: '=https://login.microsoftonline.com/{{$self["tenant"]}}/oauth2/v2.0/token',
			required: true,
		},
		{
			displayName: 'Enable Video Upload (OneDrive)',
			name: 'enableOneDrive',
			type: 'boolean',
			default: true,
			description:
				'Whether to also request access to OneDrive. Needed to upload videos to OneDrive and link them in a page. Reconnect the credential after changing this.',
		},
		{
			displayName: 'Scope',
			name: 'scope',
			type: 'hidden',
			default:
				'=offline_access User.Read Notes.ReadWrite{{$self["enableOneDrive"] ? " Files.ReadWrite" : ""}}',
		},
		{
			displayName: 'Auth URI Query Parameters',
			name: 'authQueryParameters',
			type: 'hidden',
			default: 'response_mode=query',
		},
		{
			displayName: 'Authentication',
			name: 'authentication',
			type: 'hidden',
			default: 'body',
		},
	];
}

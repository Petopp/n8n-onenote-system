import type {
	IDataObject,
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
} from 'n8n-workflow';
import { NodeConnectionTypes, NodeOperationError, sleep } from 'n8n-workflow';
import { authority, FORM_HEADERS, formBody, scopes } from './utils/auth';

// Deliberately not usable as AI tool: the node outputs a secret refresh token.
// eslint-disable-next-line @n8n/community-nodes/node-usable-as-tool
export class OneNoteLogin implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'OneNote Login Helper',
		name: 'oneNoteLogin',
		icon: { light: 'file:onenote.svg', dark: 'file:onenote.dark.svg' },
		group: ['input'],
		version: 1,
		subtitle: '={{$parameter["operation"]}}',
		description:
			'Creates the refresh token for the "OneNote Device Login API" credential (device code flow, no redirect URL needed)',
		defaults: { name: 'OneNote Login Helper' },
		inputs: [NodeConnectionTypes.Main],
		outputs: [NodeConnectionTypes.Main],
		properties: [
			{
				displayName:
					'Step 1: run "Start Login" and open the shown address. Step 2: enter the code there, then run "Finish Login" and copy the refresh token into the credential. Delete the execution afterwards - the token is a secret.',
				name: 'notice',
				type: 'notice',
				default: '',
			},
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				options: [
					{ name: 'Start Login', value: 'start', action: 'Start the device login' },
					{ name: 'Finish Login', value: 'finish', action: 'Finish the device login' },
				],
				default: 'start',
			},
			{
				displayName: 'Client ID',
				name: 'clientId',
				type: 'string',
				default: '',
				required: true,
				description: 'Application (client) ID of the Azure app registration',
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
			},
			{
				displayName: 'Device Code',
				name: 'deviceCode',
				type: 'string',
				default: '={{ $json.deviceCode }}',
				required: true,
				displayOptions: { show: { operation: ['finish'] } },
				description: 'The device code returned by "Start Login"',
			},
			{
				displayName: 'Wait for Confirmation (Seconds)',
				name: 'waitSeconds',
				type: 'number',
				typeOptions: { minValue: 0, maxValue: 600 },
				default: 120,
				displayOptions: { show: { operation: ['finish'] } },
				description: 'How long to wait until the code has been entered in the browser',
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const operation = this.getNodeParameter('operation', 0) as string;
		const clientId = this.getNodeParameter('clientId', 0) as string;
		const tenant = this.getNodeParameter('tenant', 0) as string;
		const oneDrive = this.getNodeParameter('enableOneDrive', 0) as boolean;
		const base = authority(tenant);
		const post = async (path: string, values: Record<string, string>): Promise<IDataObject> =>
			(await this.helpers.httpRequest({
				method: 'POST',
				url: `${base}/${path}`,
				headers: FORM_HEADERS,
				body: formBody(values),
				json: true,
				ignoreHttpStatusErrors: true,
			})) as IDataObject;

		if (operation === 'start') {
			const r = await post('devicecode', { client_id: clientId, scope: scopes(oneDrive) });
			if (r.error) {
				throw new NodeOperationError(
					this.getNode(),
					`${String(r.error_description ?? r.error)} - check the Client ID, that the app has the platform "Mobile and desktop applications" and that "Allow public client flows" is enabled.`,
				);
			}
			return [
				[
					{
						json: {
							verificationUri: r.verification_uri,
							userCode: r.user_code,
							deviceCode: r.device_code,
							expiresInSeconds: r.expires_in,
							message: r.message,
						},
						pairedItem: { item: 0 },
					},
				],
			];
		}

		const deviceCode = this.getNodeParameter('deviceCode', 0) as string;
		const waitSeconds = this.getNodeParameter('waitSeconds', 0) as number;
		const deadline = Date.now() + waitSeconds * 1000;
		let intervalMs = 5000;
		for (;;) {
			const r = await post('token', {
				client_id: clientId,
				grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
				device_code: deviceCode,
			});
			if (r.refresh_token) {
				return [
					[
						{
							json: {
								refreshToken: r.refresh_token,
								note: 'Copy this into the credential "OneNote Device Login API", then delete this execution.',
							},
							pairedItem: { item: 0 },
						},
					],
				];
			}
			if (r.error === 'authorization_pending' || r.error === 'slow_down') {
				if (r.error === 'slow_down') intervalMs += 5000;
				if (Date.now() + intervalMs > deadline) {
					throw new NodeOperationError(
						this.getNode(),
						'Not confirmed yet. Enter the code at the shown address, then run "Finish Login" again.',
					);
				}
				await sleep(intervalMs);
				continue;
			}
			throw new NodeOperationError(
				this.getNode(),
				String(r.error_description ?? r.error ?? 'Login failed'),
			);
		}
	}
}

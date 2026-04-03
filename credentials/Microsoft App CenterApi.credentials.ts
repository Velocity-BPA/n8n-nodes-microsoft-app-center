import {
	ICredentialType,
	INodeProperties,
} from 'n8n-workflow';

export class MicrosoftAppCenterApi implements ICredentialType {
	name = 'microsoftAppCenterApi';
	displayName = 'Microsoft App Center API';
	documentationUrl = 'https://docs.microsoft.com/en-us/appcenter/api/';
	properties: INodeProperties[] = [
		{
			displayName: 'API Token',
			name: 'apiToken',
			type: 'string',
			typeOptions: {
				password: true,
			},
			default: '',
			description: 'API token generated from App Center portal under Account Settings > User API tokens',
			required: true,
		},
		{
			displayName: 'Base URL',
			name: 'baseUrl',
			type: 'string',
			default: 'https://api.appcenter.ms/v0.1',
			description: 'Base URL for Microsoft App Center API',
			required: true,
		},
	];
}
/**
 * Copyright (c) 2026 Velocity BPA
 * 
 * Licensed under the Business Source License 1.1 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *     https://github.com/VelocityBPA/n8n-nodes-microsoftappcenter/blob/main/LICENSE
 * 
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import {
  IExecuteFunctions,
  INodeExecutionData,
  INodeType,
  INodeTypeDescription,
  NodeOperationError,
  NodeApiError,
} from 'n8n-workflow';

export class MicrosoftAppCenter implements INodeType {
  description: INodeTypeDescription = {
    displayName: 'Microsoft App Center',
    name: 'microsoftappcenter',
    icon: 'file:microsoftappcenter.svg',
    group: ['transform'],
    version: 1,
    subtitle: '={{$parameter["operation"] + ": " + $parameter["resource"]}}',
    description: 'Interact with the Microsoft App Center API',
    defaults: {
      name: 'Microsoft App Center',
    },
    inputs: ['main'],
    outputs: ['main'],
    credentials: [
      {
        name: 'microsoftappcenterApi',
        required: true,
      },
    ],
    properties: [
      {
        displayName: 'Resource',
        name: 'resource',
        type: 'options',
        noDataExpression: true,
        options: [
          {
            name: 'App',
            value: 'app',
          },
          {
            name: 'Build',
            value: 'build',
          },
          {
            name: 'Release',
            value: 'release',
          },
          {
            name: 'Crash',
            value: 'crash',
          },
          {
            name: 'Analytics',
            value: 'analytics',
          },
          {
            name: 'CodePush',
            value: 'codePush',
          },
          {
            name: 'Distribution Group',
            value: 'distributionGroup',
          }
        ],
        default: 'app',
      },
{
  displayName: 'Operation',
  name: 'operation',
  type: 'options',
  noDataExpression: true,
  displayOptions: { show: { resource: ['app'] } },
  options: [
    { name: 'Get All', value: 'getAll', description: 'List all apps accessible to the user', action: 'Get all apps' },
    { name: 'Create', value: 'create', description: 'Create a new app', action: 'Create an app' },
    { name: 'Get', value: 'get', description: 'Get specific app details', action: 'Get an app' },
    { name: 'Update', value: 'update', description: 'Update app information', action: 'Update an app' },
    { name: 'Delete', value: 'delete', description: 'Delete an app', action: 'Delete an app' }
  ],
  default: 'getAll',
},
{
  displayName: 'Operation',
  name: 'operation',
  type: 'options',
  noDataExpression: true,
  displayOptions: { show: { resource: ['build'] } },
  options: [
    { name: 'Get All Builds', value: 'getAll', description: 'List builds for an app', action: 'Get all builds' },
    { name: 'Create Build', value: 'create', description: 'Trigger a new build', action: 'Create a build' },
    { name: 'Get Build', value: 'get', description: 'Get build details', action: 'Get a build' },
    { name: 'Get Build Logs', value: 'getLogs', description: 'Get build logs', action: 'Get build logs' },
    { name: 'Get Branches', value: 'getBranches', description: 'List configured branches', action: 'Get branches' },
  ],
  default: 'getAll',
},
{
  displayName: 'Operation',
  name: 'operation',
  type: 'options',
  noDataExpression: true,
  displayOptions: { show: { resource: ['release'] } },
  options: [
    { name: 'Get All', value: 'getAll', description: 'List releases for an app', action: 'Get all releases' },
    { name: 'Create', value: 'create', description: 'Create release from upload', action: 'Create a release' },
    { name: 'Get', value: 'get', description: 'Get release details', action: 'Get a release' },
    { name: 'Update', value: 'update', description: 'Update release information', action: 'Update a release' },
    { name: 'Delete', value: 'delete', description: 'Delete a release', action: 'Delete a release' }
  ],
  default: 'getAll',
},
{
  displayName: 'Operation',
  name: 'operation',
  type: 'options',
  noDataExpression: true,
  displayOptions: { show: { resource: ['crash'] } },
  options: [
    { name: 'Get All Crash Groups', value: 'getAll', description: 'List crash groups for an app', action: 'Get all crash groups' },
    { name: 'Get Crash Group', value: 'get', description: 'Get details of a specific crash group', action: 'Get crash group details' },
    { name: 'Update Crash Group', value: 'update', description: 'Update crash group status', action: 'Update crash group status' },
    { name: 'Get All Crashes', value: 'getCrashes', description: 'List individual crashes', action: 'Get all crashes' },
    { name: 'Get Crash', value: 'getCrash', description: 'Get specific crash details', action: 'Get crash details' },
  ],
  default: 'getAll',
},
{
  displayName: 'Operation',
  name: 'operation',
  type: 'options',
  noDataExpression: true,
  displayOptions: { show: { resource: ['analytics'] } },
  options: [
    { name: 'Get Events', value: 'getEvents', description: 'Get analytics events', action: 'Get events' },
    { name: 'Get Sessions', value: 'getSessions', description: 'Get session analytics', action: 'Get sessions' },
    { name: 'Get Audiences', value: 'getAudiences', description: 'Get audience analytics', action: 'Get audiences' },
    { name: 'Get Versions', value: 'getVersions', description: 'Get version analytics', action: 'Get versions' },
    { name: 'Get Distribution', value: 'getDistribution', description: 'Get distribution analytics', action: 'Get distribution' },
  ],
  default: 'getEvents',
},
{
	displayName: 'Operation',
	name: 'operation',
	type: 'options',
	noDataExpression: true,
	displayOptions: { show: { resource: ['codePush'] } },
	options: [
		{ name: 'Get All Deployments', value: 'getAll', description: 'List all CodePush deployments for an app', action: 'Get all deployments' },
		{ name: 'Create Deployment', value: 'create', description: 'Create a new CodePush deployment', action: 'Create deployment' },
		{ name: 'Get Deployment', value: 'get', description: 'Get details of a specific deployment', action: 'Get deployment' },
		{ name: 'Update Deployment', value: 'update', description: 'Update an existing deployment', action: 'Update deployment' },
		{ name: 'Delete Deployment', value: 'delete', description: 'Delete a deployment', action: 'Delete deployment' }
	],
	default: 'getAll',
},
{
	displayName: 'Operation',
	name: 'operation',
	type: 'options',
	noDataExpression: true,
	displayOptions: {
		show: {
			resource: ['distributionGroup'],
		},
	},
	options: [
		{
			name: 'Create',
			value: 'create',
			description: 'Create a distribution group',
			action: 'Create a distribution group',
		},
		{
			name: 'Delete',
			value: 'delete',
			description: 'Delete a distribution group',
			action: 'Delete a distribution group',
		},
		{
			name: 'Get',
			value: 'get',
			description: 'Get distribution group details',
			action: 'Get a distribution group',
		},
		{
			name: 'Get All',
			value: 'getAll',
			description: 'List all distribution groups',
			action: 'Get all distribution groups',
		},
		{
			name: 'Update',
			value: 'update',
			description: 'Update a distribution group',
			action: 'Update a distribution group',
		},
	],
	default: 'getAll',
},
{
  displayName: 'Display Name',
  name: 'display_name',
  type: 'string',
  required: true,
  displayOptions: { show: { resource: ['app'], operation: ['create'] } },
  default: '',
  description: 'The display name of the app'
},
{
  displayName: 'Name',
  name: 'name',
  type: 'string',
  required: true,
  displayOptions: { show: { resource: ['app'], operation: ['create'] } },
  default: '',
  description: 'The name of the app'
},
{
  displayName: 'OS',
  name: 'os',
  type: 'options',
  required: true,
  displayOptions: { show: { resource: ['app'], operation: ['create'] } },
  options: [
    { name: 'Android', value: 'Android' },
    { name: 'iOS', value: 'iOS' },
    { name: 'Windows', value: 'Windows' },
    { name: 'macOS', value: 'macOS' }
  ],
  default: 'Android',
  description: 'The OS the app will be running on'
},
{
  displayName: 'Platform',
  name: 'platform',
  type: 'options',
  required: true,
  displayOptions: { show: { resource: ['app'], operation: ['create'] } },
  options: [
    { name: 'Java', value: 'Java' },
    { name: 'Objective-C-Swift', value: 'Objective-C-Swift' },
    { name: 'React-Native', value: 'React-Native' },
    { name: 'Unity', value: 'Unity' },
    { name: 'Xamarin', value: 'Xamarin' },
    { name: 'Cordova', value: 'Cordova' }
  ],
  default: 'Java',
  description: 'The platform of the app'
},
{
  displayName: 'Release Type',
  name: 'release_type',
  type: 'options',
  required: false,
  displayOptions: { show: { resource: ['app'], operation: ['create'] } },
  options: [
    { name: 'Alpha', value: 'alpha' },
    { name: 'Beta', value: 'beta' },
    { name: 'Production', value: 'production' },
    { name: 'Store', value: 'store' }
  ],
  default: 'beta',
  description: 'The release type of the app'
},
{
  displayName: 'Owner Name',
  name: 'owner_name',
  type: 'string',
  required: true,
  displayOptions: { show: { resource: ['app'], operation: ['get', 'update', 'delete'] } },
  default: '',
  description: 'The name of the owner'
},
{
  displayName: 'App Name',
  name: 'app_name',
  type: 'string',
  required: true,
  displayOptions: { show: { resource: ['app'], operation: ['get', 'update', 'delete'] } },
  default: '',
  description: 'The name of the app'
},
{
  displayName: 'Display Name',
  name: 'display_name',
  type: 'string',
  required: false,
  displayOptions: { show: { resource: ['app'], operation: ['update'] } },
  default: '',
  description: 'The display name of the app'
},
{
  displayName: 'Description',
  name: 'description',
  type: 'string',
  required: false,
  displayOptions: { show: { resource: ['app'], operation: ['update'] } },
  default: '',
  description: 'The description of the app'
},
{
  displayName: 'Owner Name',
  name: 'ownerName',
  type: 'string',
  required: true,
  displayOptions: { show: { resource: ['build'], operation: ['getAll', 'create', 'get', 'getLogs', 'getBranches'] } },
  default: '',
  description: 'The name of the app owner',
},
{
  displayName: 'App Name',
  name: 'appName',
  type: 'string',
  required: true,
  displayOptions: { show: { resource: ['build'], operation: ['getAll', 'create', 'get', 'getLogs', 'getBranches'] } },
  default: '',
  description: 'The name of the app',
},
{
  displayName: 'Branch',
  name: 'branch',
  type: 'string',
  required: false,
  displayOptions: { show: { resource: ['build'], operation: ['getAll'] } },
  default: '',
  description: 'Filter builds by branch name',
},
{
  displayName: 'Build ID',
  name: 'buildId',
  type: 'string',
  required: true,
  displayOptions: { show: { resource: ['build'], operation: ['get', 'getLogs'] } },
  default: '',
  description: 'The ID of the build',
},
{
  displayName: 'Source Version',
  name: 'sourceVersion',
  type: 'string',
  required: false,
  displayOptions: { show: { resource: ['build'], operation: ['create'] } },
  default: '',
  description: 'The commit hash or branch to build from',
},
{
  displayName: 'Debug',
  name: 'debug',
  type: 'boolean',
  required: false,
  displayOptions: { show: { resource: ['build'], operation: ['create'] } },
  default: false,
  description: 'Whether to enable debug mode for the build',
},
{
  displayName: 'Owner Name',
  name: 'ownerName',
  type: 'string',
  required: true,
  displayOptions: { show: { resource: ['release'], operation: ['getAll', 'create', 'get', 'update', 'delete'] } },
  default: '',
  description: 'The name of the owner',
},
{
  displayName: 'App Name',
  name: 'appName',
  type: 'string',
  required: true,
  displayOptions: { show: { resource: ['release'], operation: ['getAll', 'create', 'get', 'update', 'delete'] } },
  default: '',
  description: 'The name of the application',
},
{
  displayName: 'Published Only',
  name: 'publishedOnly',
  type: 'boolean',
  displayOptions: { show: { resource: ['release'], operation: ['getAll'] } },
  default: false,
  description: 'Return only published releases',
},
{
  displayName: 'Release ID',
  name: 'releaseId',
  type: 'string',
  required: true,
  displayOptions: { show: { resource: ['release'], operation: ['create', 'get', 'update', 'delete'] } },
  default: '',
  description: 'The ID of the release',
},
{
  displayName: 'Enabled',
  name: 'enabled',
  type: 'boolean',
  displayOptions: { show: { resource: ['release'], operation: ['update'] } },
  default: true,
  description: 'Whether the release is enabled',
},
{
  displayName: 'Owner Name',
  name: 'ownerName',
  type: 'string',
  required: true,
  displayOptions: { show: { resource: ['crash'], operation: ['getAll', 'get', 'update', 'getCrashes', 'getCrash'] } },
  default: '',
  description: 'The name of the owner',
},
{
  displayName: 'App Name',
  name: 'appName',
  type: 'string',
  required: true,
  displayOptions: { show: { resource: ['crash'], operation: ['getAll', 'get', 'update', 'getCrashes', 'getCrash'] } },
  default: '',
  description: 'The name of the application',
},
{
  displayName: 'Crash Group ID',
  name: 'crashGroupId',
  type: 'string',
  required: true,
  displayOptions: { show: { resource: ['crash'], operation: ['get', 'update'] } },
  default: '',
  description: 'The ID of the crash group',
},
{
  displayName: 'Crash ID',
  name: 'crashId',
  type: 'string',
  required: true,
  displayOptions: { show: { resource: ['crash'], operation: ['getCrash'] } },
  default: '',
  description: 'The ID of the crash',
},
{
  displayName: 'Status',
  name: 'status',
  type: 'options',
  required: true,
  displayOptions: { show: { resource: ['crash'], operation: ['update'] } },
  options: [
    { name: 'Open', value: 'open' },
    { name: 'Closed', value: 'closed' },
    { name: 'Ignored', value: 'ignored' },
  ],
  default: 'open',
  description: 'The status to set for the crash group',
},
{
  displayName: 'Version',
  name: 'version',
  type: 'string',
  displayOptions: { show: { resource: ['crash'], operation: ['getAll'] } },
  default: '',
  description: 'Filter by app version',
},
{
  displayName: 'Start Time',
  name: 'startTime',
  type: 'dateTime',
  displayOptions: { show: { resource: ['crash'], operation: ['getAll', 'getCrashes'] } },
  default: '',
  description: 'Start time for filtering results',
},
{
  displayName: 'End Time',
  name: 'endTime',
  type: 'dateTime',
  displayOptions: { show: { resource: ['crash'], operation: ['getAll', 'getCrashes'] } },
  default: '',
  description: 'End time for filtering results',
},
{
  displayName: 'Crash Group ID',
  name: 'crashGroupIdFilter',
  type: 'string',
  displayOptions: { show: { resource: ['crash'], operation: ['getCrashes'] } },
  default: '',
  description: 'Filter crashes by crash group ID',
},
{
  displayName: 'Owner Name',
  name: 'ownerName',
  type: 'string',
  required: true,
  displayOptions: { show: { resource: ['analytics'], operation: ['getEvents', 'getSessions', 'getAudiences', 'getVersions', 'getDistribution'] } },
  default: '',
  description: 'The name of the owner',
},
{
  displayName: 'App Name',
  name: 'appName',
  type: 'string',
  required: true,
  displayOptions: { show: { resource: ['analytics'], operation: ['getEvents', 'getSessions', 'getAudiences', 'getVersions', 'getDistribution'] } },
  default: '',
  description: 'The name of the app',
},
{
  displayName: 'Start Date',
  name: 'start',
  type: 'dateTime',
  required: true,
  displayOptions: { show: { resource: ['analytics'], operation: ['getEvents', 'getSessions', 'getAudiences', 'getVersions', 'getDistribution'] } },
  default: '',
  description: 'Start date for analytics data',
},
{
  displayName: 'End Date',
  name: 'end',
  type: 'dateTime',
  required: true,
  displayOptions: { show: { resource: ['analytics'], operation: ['getEvents', 'getSessions', 'getAudiences', 'getVersions', 'getDistribution'] } },
  default: '',
  description: 'End date for analytics data',
},
{
  displayName: 'Versions',
  name: 'versions',
  type: 'string',
  required: false,
  displayOptions: { show: { resource: ['analytics'], operation: ['getEvents', 'getSessions'] } },
  default: '',
  description: 'Comma-separated list of version names to filter by',
},
{
  displayName: 'Top',
  name: 'top',
  type: 'number',
  required: false,
  displayOptions: { show: { resource: ['analytics'], operation: ['getVersions'] } },
  default: 10,
  description: 'Number of top entries to return',
},
{
	displayName: 'Owner Name',
	name: 'ownerName',
	type: 'string',
	required: true,
	displayOptions: {
		show: {
			resource: ['codePush'],
			operation: ['getAll', 'create', 'get', 'update', 'delete']
		}
	},
	default: '',
	placeholder: 'microsoft',
	description: 'The name of the app owner'
},
{
	displayName: 'App Name',
	name: 'appName',
	type: 'string',
	required: true,
	displayOptions: {
		show: {
			resource: ['codePush'],
			operation: ['getAll', 'create', 'get', 'update', 'delete']
		}
	},
	default: '',
	placeholder: 'MyApp-iOS',
	description: 'The name of the app'
},
{
	displayName: 'Deployment Name',
	name: 'deploymentName',
	type: 'string',
	required: true,
	displayOptions: {
		show: {
			resource: ['codePush'],
			operation: ['get', 'update', 'delete']
		}
	},
	default: '',
	placeholder: 'Production',
	description: 'The name of the deployment'
},
{
	displayName: 'Name',
	name: 'name',
	type: 'string',
	required: true,
	displayOptions: {
		show: {
			resource: ['codePush'],
			operation: ['create']
		}
	},
	default: '',
	placeholder: 'Production',
	description: 'Name for the new deployment'
},
{
	displayName: 'New Name',
	name: 'newName',
	type: 'string',
	required: false,
	displayOptions: {
		show: {
			resource: ['codePush'],
			operation: ['update']
		}
	},
	default: '',
	placeholder: 'Staging',
	description: 'New name for the deployment'
},
{
	displayName: 'Owner Name',
	name: 'ownerName',
	type: 'string',
	required: true,
	displayOptions: {
		show: {
			resource: ['distributionGroup'],
			operation: ['create', 'delete', 'get', 'getAll', 'update'],
		},
	},
	default: '',
	description: 'The name of the owner',
},
{
	displayName: 'App Name',
	name: 'appName',
	type: 'string',
	required: true,
	displayOptions: {
		show: {
			resource: ['distributionGroup'],
			operation: ['create', 'delete', 'get', 'getAll', 'update'],
		},
	},
	default: '',
	description: 'The name of the application',
},
{
	displayName: 'Distribution Group Name',
	name: 'distributionGroupName',
	type: 'string',
	required: true,
	displayOptions: {
		show: {
			resource: ['distributionGroup'],
			operation: ['delete', 'get', 'update'],
		},
	},
	default: '',
	description: 'The name of the distribution group',
},
{
	displayName: 'Name',
	name: 'name',
	type: 'string',
	required: true,
	displayOptions: {
		show: {
			resource: ['distributionGroup'],
			operation: ['create'],
		},
	},
	default: '',
	description: 'The name of the distribution group',
},
{
	displayName: 'Name',
	name: 'name',
	type: 'string',
	required: false,
	displayOptions: {
		show: {
			resource: ['distributionGroup'],
			operation: ['update'],
		},
	},
	default: '',
	description: 'The new name of the distribution group',
},
{
	displayName: 'Is Public',
	name: 'isPublic',
	type: 'boolean',
	displayOptions: {
		show: {
			resource: ['distributionGroup'],
			operation: ['create', 'update'],
		},
	},
	default: false,
	description: 'Whether the distribution group is public',
},
    ],
  };

  async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
    const items = this.getInputData();
    const resource = this.getNodeParameter('resource', 0) as string;

    switch (resource) {
      case 'app':
        return [await executeAppOperations.call(this, items)];
      case 'build':
        return [await executeBuildOperations.call(this, items)];
      case 'release':
        return [await executeReleaseOperations.call(this, items)];
      case 'crash':
        return [await executeCrashOperations.call(this, items)];
      case 'analytics':
        return [await executeAnalyticsOperations.call(this, items)];
      case 'codePush':
        return [await executeCodePushOperations.call(this, items)];
      case 'distributionGroup':
        return [await executeDistributionGroupOperations.call(this, items)];
      default:
        throw new NodeOperationError(this.getNode(), `The resource "${resource}" is not supported`);
    }
  }
}

// ============================================================
// Resource Handler Functions
// ============================================================

async function executeAppOperations(
  this: IExecuteFunctions,
  items: INodeExecutionData[],
): Promise<INodeExecutionData[]> {
  const returnData: INodeExecutionData[] = [];
  const operation = this.getNodeParameter('operation', 0) as string;
  const credentials = await this.getCredentials('microsoftappcenterApi') as any;

  for (let i = 0; i < items.length; i++) {
    try {
      let result: any;

      switch (operation) {
        case 'getAll': {
          const options: any = {
            method: 'GET',
            url: `${credentials.baseUrl}/apps`,
            headers: {
              'X-API-Token': credentials.apiKey,
              'Content-Type': 'application/json',
            },
            json: true,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }

        case 'create': {
          const displayName = this.getNodeParameter('display_name', i) as string;
          const name = this.getNodeParameter('name', i) as string;
          const os = this.getNodeParameter('os', i) as string;
          const platform = this.getNodeParameter('platform', i) as string;
          const releaseType = this.getNodeParameter('release_type', i) as string;

          const body: any = {
            display_name: displayName,
            name: name,
            os: os,
            platform: platform,
          };

          if (releaseType) {
            body.release_type = releaseType;
          }

          const options: any = {
            method: 'POST',
            url: `${credentials.baseUrl}/apps`,
            headers: {
              'X-API-Token': credentials.apiKey,
              'Content-Type': 'application/json',
            },
            body: body,
            json: true,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }

        case 'get': {
          const ownerName = this.getNodeParameter('owner_name', i) as string;
          const appName = this.getNodeParameter('app_name', i) as string;

          const options: any = {
            method: 'GET',
            url: `${credentials.baseUrl}/apps/${ownerName}/${appName}`,
            headers: {
              'X-API-Token': credentials.apiKey,
              'Content-Type': 'application/json',
            },
            json: true,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }

        case 'update': {
          const ownerName = this.getNodeParameter('owner_name', i) as string;
          const appName = this.getNodeParameter('app_name', i) as string;
          const displayName = this.getNodeParameter('display_name', i) as string;
          const description = this.getNodeParameter('description', i) as string;

          const body: any = {};
          if (displayName) {
            body.display_name = displayName;
          }
          if (description) {
            body.description = description;
          }

          const options: any = {
            method: 'PATCH',
            url: `${credentials.baseUrl}/apps/${ownerName}/${appName}`,
            headers: {
              'X-API-Token': credentials.apiKey,
              'Content-Type': 'application/json',
            },
            body: body,
            json: true,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }

        case 'delete': {
          const ownerName = this.getNodeParameter('owner_name', i) as string;
          const appName = this.getNodeParameter('app_name', i) as string;

          const options: any = {
            method: 'DELETE',
            url: `${credentials.baseUrl}/apps/${ownerName}/${appName}`,
            headers: {
              'X-API-Token': credentials.apiKey,
              'Content-Type': 'application/json',
            },
            json: true,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }

        default:
          throw new NodeOperationError(this.getNode(), `Unknown operation: ${operation}`);
      }

      returnData.push({ json: result, pairedItem: { item: i } });
    } catch (error: any) {
      if (this.continueOnFail()) {
        returnData.push({ json: { error: error.message }, pairedItem: { item: i } });
      } else {
        throw error;
      }
    }
  }

  return returnData;
}

async function executeBuildOperations(
  this: IExecuteFunctions,
  items: INodeExecutionData[],
): Promise<INodeExecutionData[]> {
  const returnData: INodeExecutionData[] = [];
  const operation = this.getNodeParameter('operation', 0) as string;
  const credentials = await this.getCredentials('microsoftappcenterApi') as any;

  for (let i = 0; i < items.length; i++) {
    try {
      let result: any;
      const ownerName = this.getNodeParameter('ownerName', i) as string;
      const appName = this.getNodeParameter('appName', i) as string;

      switch (operation) {
        case 'getAll': {
          const branch = this.getNodeParameter('branch', i) as string;
          let endpoint = `/apps/${ownerName}/${appName}/builds`;
          if (branch) {
            endpoint += `?branch=${encodeURIComponent(branch)}`;
          }
          
          const options: any = {
            method: 'GET',
            url: credentials.baseUrl + endpoint,
            headers: {
              'X-API-Token': credentials.apiKey,
              'Content-Type': 'application/json',
            },
            json: true,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }

        case 'create': {
          const sourceVersion = this.getNodeParameter('sourceVersion', i) as string;
          const debug = this.getNodeParameter('debug', i) as boolean;
          
          const body: any = {};
          if (sourceVersion) body.sourceVersion = sourceVersion;
          if (debug !== undefined) body.debug = debug;

          const options: any = {
            method: 'POST',
            url: credentials.baseUrl + `/apps/${ownerName}/${appName}/builds`,
            headers: {
              'X-API-Token': credentials.apiKey,
              'Content-Type': 'application/json',
            },
            body,
            json: true,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }

        case 'get': {
          const buildId = this.getNodeParameter('buildId', i) as string;
          
          const options: any = {
            method: 'GET',
            url: credentials.baseUrl + `/apps/${ownerName}/${appName}/builds/${buildId}`,
            headers: {
              'X-API-Token': credentials.apiKey,
              'Content-Type': 'application/json',
            },
            json: true,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }

        case 'getLogs': {
          const buildId = this.getNodeParameter('buildId', i) as string;
          
          const options: any = {
            method: 'GET',
            url: credentials.baseUrl + `/apps/${ownerName}/${appName}/builds/${buildId}/logs`,
            headers: {
              'X-API-Token': credentials.apiKey,
              'Content-Type': 'application/json',
            },
            json: true,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }

        case 'getBranches': {
          const options: any = {
            method: 'GET',
            url: credentials.baseUrl + `/apps/${ownerName}/${appName}/branches`,
            headers: {
              'X-API-Token': credentials.apiKey,
              'Content-Type': 'application/json',
            },
            json: true,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }

        default:
          throw new NodeOperationError(this.getNode(), 'Unknown operation: ' + operation);
      }

      returnData.push({ json: result, pairedItem: { item: i } });
    } catch (error: any) {
      if (this.continueOnFail()) {
        returnData.push({ json: { error: error.message }, pairedItem: { item: i } });
      } else {
        throw error;
      }
    }
  }

  return returnData;
}

async function executeReleaseOperations(
  this: IExecuteFunctions,
  items: INodeExecutionData[],
): Promise<INodeExecutionData[]> {
  const returnData: INodeExecutionData[] = [];
  const operation = this.getNodeParameter('operation', 0) as string;
  const credentials = await this.getCredentials('microsoftappcenterApi') as any;

  for (let i = 0; i < items.length; i++) {
    try {
      let result: any;

      const ownerName = this.getNodeParameter('ownerName', i) as string;
      const appName = this.getNodeParameter('appName', i) as string;

      const baseOptions = {
        headers: {
          'X-API-Token': credentials.apiKey,
          'Content-Type': 'application/json',
        },
        json: true,
      };

      switch (operation) {
        case 'getAll': {
          const publishedOnly = this.getNodeParameter('publishedOnly', i) as boolean;
          let url = `${credentials.baseUrl || 'https://api.appcenter.ms/v0.1'}/apps/${ownerName}/${appName}/releases`;
          if (publishedOnly) {
            url += '?published_only=true';
          }
          const options: any = {
            ...baseOptions,
            method: 'GET',
            url,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }
        case 'create': {
          const releaseId = this.getNodeParameter('releaseId', i) as string;
          const options: any = {
            ...baseOptions,
            method: 'POST',
            url: `${credentials.baseUrl || 'https://api.appcenter.ms/v0.1'}/apps/${ownerName}/${appName}/releases/${releaseId}`,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }
        case 'get': {
          const releaseId = this.getNodeParameter('releaseId', i) as string;
          const options: any = {
            ...baseOptions,
            method: 'GET',
            url: `${credentials.baseUrl || 'https://api.appcenter.ms/v0.1'}/apps/${ownerName}/${appName}/releases/${releaseId}`,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }
        case 'update': {
          const releaseId = this.getNodeParameter('releaseId', i) as string;
          const enabled = this.getNodeParameter('enabled', i) as boolean;
          const options: any = {
            ...baseOptions,
            method: 'PATCH',
            url: `${credentials.baseUrl || 'https://api.appcenter.ms/v0.1'}/apps/${ownerName}/${appName}/releases/${releaseId}`,
            body: {
              enabled,
            },
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }
        case 'delete': {
          const releaseId = this.getNodeParameter('releaseId', i) as string;
          const options: any = {
            ...baseOptions,
            method: 'DELETE',
            url: `${credentials.baseUrl || 'https://api.appcenter.ms/v0.1'}/apps/${ownerName}/${appName}/releases/${releaseId}`,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }
        default:
          throw new NodeOperationError(this.getNode(), `Unknown operation: ${operation}`);
      }

      returnData.push({ json: result, pairedItem: { item: i } });
    } catch (error: any) {
      if (this.continueOnFail()) {
        returnData.push({ json: { error: error.message }, pairedItem: { item: i } });
      } else {
        throw error;
      }
    }
  }

  return returnData;
}

async function executeCrashOperations(
  this: IExecuteFunctions,
  items: INodeExecutionData[],
): Promise<INodeExecutionData[]> {
  const returnData: INodeExecutionData[] = [];
  const operation = this.getNodeParameter('operation', 0) as string;
  const credentials = await this.getCredentials('microsoftappcenterApi') as any;

  for (let i = 0; i < items.length; i++) {
    try {
      let result: any;
      const ownerName = this.getNodeParameter('ownerName', i) as string;
      const appName = this.getNodeParameter('appName', i) as string;

      const baseUrl = credentials.baseUrl || 'https://api.appcenter.ms/v0.1';
      const headers = {
        'X-API-Token': credentials.apiKey,
        'Content-Type': 'application/json',
      };

      switch (operation) {
        case 'getAll': {
          const qs: any = {};
          const version = this.getNodeParameter('version', i) as string;
          const startTime = this.getNodeParameter('startTime', i) as string;
          const endTime = this.getNodeParameter('endTime', i) as string;

          if (version) qs.version = version;
          if (startTime) qs.start_time = startTime;
          if (endTime) qs.end_time = endTime;

          const options: any = {
            method: 'GET',
            url: `${baseUrl}/apps/${ownerName}/${appName}/crash_groups`,
            headers,
            qs,
            json: true,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }
        case 'get': {
          const crashGroupId = this.getNodeParameter('crashGroupId', i) as string;
          const options: any = {
            method: 'GET',
            url: `${baseUrl}/apps/${ownerName}/${appName}/crash_groups/${crashGroupId}`,
            headers,
            json: true,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }
        case 'update': {
          const crashGroupId = this.getNodeParameter('crashGroupId', i) as string;
          const status = this.getNodeParameter('status', i) as string;
          const options: any = {
            method: 'PATCH',
            url: `${baseUrl}/apps/${ownerName}/${appName}/crash_groups/${crashGroupId}`,
            headers,
            body: { status },
            json: true,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }
        case 'getCrashes': {
          const qs: any = {};
          const crashGroupIdFilter = this.getNodeParameter('crashGroupIdFilter', i) as string;
          const startTime = this.getNodeParameter('startTime', i) as string;
          const endTime = this.getNodeParameter('endTime', i) as string;

          if (crashGroupIdFilter) qs.crash_group_id = crashGroupIdFilter;
          if (startTime) qs.start_time = startTime;
          if (endTime) qs.end_time = endTime;

          const options: any = {
            method: 'GET',
            url: `${baseUrl}/apps/${ownerName}/${appName}/crashes`,
            headers,
            qs,
            json: true,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }
        case 'getCrash': {
          const crashId = this.getNodeParameter('crashId', i) as string;
          const options: any = {
            method: 'GET',
            url: `${baseUrl}/apps/${ownerName}/${appName}/crashes/${crashId}`,
            headers,
            json: true,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }
        default:
          throw new NodeOperationError(this.getNode(), `Unknown operation: ${operation}`);
      }

      returnData.push({ json: result, pairedItem: { item: i } });
    } catch (error: any) {
      if (this.continueOnFail()) {
        returnData.push({ json: { error: error.message }, pairedItem: { item: i } });
      } else {
        throw error;
      }
    }
  }
  return returnData;
}

async function executeAnalyticsOperations(
  this: IExecuteFunctions,
  items: INodeExecutionData[],
): Promise<INodeExecutionData[]> {
  const returnData: INodeExecutionData[] = [];
  const operation = this.getNodeParameter('operation', 0) as string;
  const credentials = await this.getCredentials('microsoftappcenterApi') as any;

  for (let i = 0; i < items.length; i++) {
    try {
      let result: any;
      const ownerName = this.getNodeParameter('ownerName', i) as string;
      const appName = this.getNodeParameter('appName', i) as string;
      const start = this.getNodeParameter('start', i) as string;
      const end = this.getNodeParameter('end', i) as string;

      const baseUrl = credentials.baseUrl || 'https://api.appcenter.ms/v0.1';
      const headers: any = {
        'X-API-Token': credentials.apiKey,
        'Content-Type': 'application/json',
      };

      switch (operation) {
        case 'getEvents': {
          const versions = this.getNodeParameter('versions', i, '') as string;
          let url = `${baseUrl}/apps/${ownerName}/${appName}/analytics/events?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
          if (versions) {
            url += `&versions=${encodeURIComponent(versions)}`;
          }
          
          const options: any = {
            method: 'GET',
            url,
            headers,
            json: true,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }
        
        case 'getSessions': {
          const versions = this.getNodeParameter('versions', i, '') as string;
          let url = `${baseUrl}/apps/${ownerName}/${appName}/analytics/sessions?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
          if (versions) {
            url += `&versions=${encodeURIComponent(versions)}`;
          }
          
          const options: any = {
            method: 'GET',
            url,
            headers,
            json: true,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }
        
        case 'getAudiences': {
          const url = `${baseUrl}/apps/${ownerName}/${appName}/analytics/audiences?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
          
          const options: any = {
            method: 'GET',
            url,
            headers,
            json: true,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }
        
        case 'getVersions': {
          const top = this.getNodeParameter('top', i, 10) as number;
          const url = `${baseUrl}/apps/${ownerName}/${appName}/analytics/versions?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&top=${top}`;
          
          const options: any = {
            method: 'GET',
            url,
            headers,
            json: true,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }
        
        case 'getDistribution': {
          const url = `${baseUrl}/apps/${ownerName}/${appName}/analytics/distribution?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
          
          const options: any = {
            method: 'GET',
            url,
            headers,
            json: true,
          };
          result = await this.helpers.httpRequest(options) as any;
          break;
        }
        
        default:
          throw new NodeOperationError(this.getNode(), `Unknown operation: ${operation}`);
      }

      returnData.push({ json: result, pairedItem: { item: i } });
    } catch (error: any) {
      if (this.continueOnFail()) {
        returnData.push({ json: { error: error.message }, pairedItem: { item: i } });
      } else {
        throw error;
      }
    }
  }
  return returnData;
}

async function executeCodePushOperations(
	this: IExecuteFunctions,
	items: INodeExecutionData[],
): Promise<INodeExecutionData[]> {
	const returnData: INodeExecutionData[] = [];
	const operation = this.getNodeParameter('operation', 0) as string;
	const credentials = await this.getCredentials('microsoftappcenterApi') as any;

	for (let i = 0; i < items.length; i++) {
		try {
			let result: any;
			const ownerName = this.getNodeParameter('ownerName', i) as string;
			const appName = this.getNodeParameter('appName', i) as string;

			const baseOptions: any = {
				headers: {
					'X-API-Token': credentials.apiKey,
					'Content-Type': 'application/json',
					'Accept': 'application/json'
				},
				json: true
			};

			switch (operation) {
				case 'getAll': {
					const options: any = {
						...baseOptions,
						method: 'GET',
						url: `${credentials.baseUrl || 'https://api.appcenter.ms/v0.1'}/apps/${ownerName}/${appName}/deployments`
					};
					result = await this.helpers.httpRequest(options) as any;
					break;
				}

				case 'create': {
					const name = this.getNodeParameter('name', i) as string;
					const options: any = {
						...baseOptions,
						method: 'POST',
						url: `${credentials.baseUrl || 'https://api.appcenter.ms/v0.1'}/apps/${ownerName}/${appName}/deployments`,
						body: { name }
					};
					result = await this.helpers.httpRequest(options) as any;
					break;
				}

				case 'get': {
					const deploymentName = this.getNodeParameter('deploymentName', i) as string;
					const options: any = {
						...baseOptions,
						method: 'GET',
						url: `${credentials.baseUrl || 'https://api.appcenter.ms/v0.1'}/apps/${ownerName}/${appName}/deployments/${deploymentName}`
					};
					result = await this.helpers.httpRequest(options) as any;
					break;
				}

				case 'update': {
					const deploymentName = this.getNodeParameter('deploymentName', i) as string;
					const newName = this.getNodeParameter('newName', i) as string;
					const options: any = {
						...baseOptions,
						method: 'PATCH',
						url: `${credentials.baseUrl || 'https://api.appcenter.ms/v0.1'}/apps/${ownerName}/${appName}/deployments/${deploymentName}`,
						body: { name: newName }
					};
					result = await this.helpers.httpRequest(options) as any;
					break;
				}

				case 'delete': {
					const deploymentName = this.getNodeParameter('deploymentName', i) as string;
					const options: any = {
						...baseOptions,
						method: 'DELETE',
						url: `${credentials.baseUrl || 'https://api.appcenter.ms/v0.1'}/apps/${ownerName}/${appName}/deployments/${deploymentName}`
					};
					result = await this.helpers.httpRequest(options) as any;
					if (!result) {
						result = { success: true, message: 'Deployment deleted successfully' };
					}
					break;
				}

				default:
					throw new NodeOperationError(this.getNode(), `Unknown operation: ${operation}`);
			}

			returnData.push({
				json: result,
				pairedItem: { item: i }
			});

		} catch (error: any) {
			if (this.continueOnFail()) {
				returnData.push({
					json: { error: error.message },
					pairedItem: { item: i }
				});
			} else {
				throw error;
			}
		}
	}

	return returnData;
}

async function executeDistributionGroupOperations(
	this: IExecuteFunctions,
	items: INodeExecutionData[],
): Promise<INodeExecutionData[]> {
	const returnData: INodeExecutionData[] = [];
	const operation = this.getNodeParameter('operation', 0) as string;
	const credentials = await this.getCredentials('microsoftappcenterApi') as any;

	for (let i = 0; i < items.length; i++) {
		try {
			let result: any;
			const ownerName = this.getNodeParameter('ownerName', i) as string;
			const appName = this.getNodeParameter('appName', i) as string;

			switch (operation) {
				case 'getAll': {
					const options: any = {
						method: 'GET',
						url: `${credentials.baseUrl}/apps/${ownerName}/${appName}/distribution_groups`,
						headers: {
							'X-API-Token': credentials.apiKey,
							'Content-Type': 'application/json',
						},
						json: true,
					};
					result = await this.helpers.httpRequest(options) as any;
					break;
				}

				case 'create': {
					const name = this.getNodeParameter('name', i) as string;
					const isPublic = this.getNodeParameter('isPublic', i) as boolean;

					const body: any = {
						name,
						is_public: isPublic,
					};

					const options: any = {
						method: 'POST',
						url: `${credentials.baseUrl}/apps/${ownerName}/${appName}/distribution_groups`,
						headers: {
							'X-API-Token': credentials.apiKey,
							'Content-Type': 'application/json',
						},
						body,
						json: true,
					};
					result = await this.helpers.httpRequest(options) as any;
					break;
				}

				case 'get': {
					const distributionGroupName = this.getNodeParameter('distributionGroupName', i) as string;

					const options: any = {
						method: 'GET',
						url: `${credentials.baseUrl}/apps/${ownerName}/${appName}/distribution_groups/${distributionGroupName}`,
						headers: {
							'X-API-Token': credentials.apiKey,
							'Content-Type': 'application/json',
						},
						json: true,
					};
					result = await this.helpers.httpRequest(options) as any;
					break;
				}

				case 'update': {
					const distributionGroupName = this.getNodeParameter('distributionGroupName', i) as string;
					const name = this.getNodeParameter('name', i) as string;
					const isPublic = this.getNodeParameter('isPublic', i) as boolean;

					const body: any = {};
					if (name) {
						body.name = name;
					}
					if (isPublic !== undefined) {
						body.is_public = isPublic;
					}

					const options: any = {
						method: 'PATCH',
						url: `${credentials.baseUrl}/apps/${ownerName}/${appName}/distribution_groups/${distributionGroupName}`,
						headers: {
							'X-API-Token': credentials.apiKey,
							'Content-Type': 'application/json',
						},
						body,
						json: true,
					};
					result = await this.helpers.httpRequest(options) as any;
					break;
				}

				case 'delete': {
					const distributionGroupName = this.getNodeParameter('distributionGroupName', i) as string;

					const options: any = {
						method: 'DELETE',
						url: `${credentials.baseUrl}/apps/${ownerName}/${appName}/distribution_groups/${distributionGroupName}`,
						headers: {
							'X-API-Token': credentials.apiKey,
							'Content-Type': 'application/json',
						},
						json: true,
					};
					result = await this.helpers.httpRequest(options) as any;
					break;
				}

				default:
					throw new NodeOperationError(this.getNode(), `Unknown operation: ${operation}`);
			}

			returnData.push({
				json: result,
				pairedItem: { item: i },
			});
		} catch (error: any) {
			if (this.continueOnFail()) {
				returnData.push({
					json: { error: error.message },
					pairedItem: { item: i },
				});
			} else {
				throw error;
			}
		}
	}

	return returnData;
}

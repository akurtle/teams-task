export type AuthenticatedPrincipal = {
  tenantId: string;
  objectId?: string;
  clientId?: string;
  scopes: string[];
  roles: string[];
};

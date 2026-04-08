import { NextFunction, Request, Response } from "express";
import { createRemoteJWKSet, jwtVerify, JWTPayload } from "jose";
import { env } from "../config/env";
import { AuthenticatedPrincipal } from "./authTypes";

declare module "express-serve-static-core" {
  interface Request {
    principal?: AuthenticatedPrincipal;
  }
}

const issuer = new URL(`https://login.microsoftonline.com/${env.azureAdTenantId}/v2.0`);
const jwks = createRemoteJWKSet(new URL(`${issuer.toString()}/discovery/v2.0/keys`));

export function requireTaskApiReadAccess(
  request: Request,
  response: Response,
  next: NextFunction
) {
  void authorize(request, response, next, {
    scopes: env.apiRequiredReadScopes,
    roles: env.apiRequiredReadRoles
  });
}

export function requireTaskApiWriteAccess(
  request: Request,
  response: Response,
  next: NextFunction
) {
  void authorize(request, response, next, {
    scopes: env.apiRequiredWriteScopes,
    roles: env.apiRequiredWriteRoles
  });
}

async function authorize(
  request: Request,
  response: Response,
  next: NextFunction,
  permissionSet: { scopes: string[]; roles: string[] }
) {
  try {
    const token = extractBearerToken(request.header("authorization"));

    if (!token) {
      response.status(401).json({
        error: "Unauthorized",
        message: "A bearer token is required for task API access."
      });
      return;
    }

    const verification = await jwtVerify(token, jwks, {
      issuer: issuer.toString(),
      audience: env.apiAudience
    });

    const principal = toPrincipal(verification.payload);

    if (!hasPermission(principal, permissionSet)) {
      response.status(403).json({
        error: "Forbidden",
        message: "Token does not include the required task API permissions."
      });
      return;
    }

    request.principal = principal;
    next();
  } catch (error) {
    response.status(401).json({
      error: "Unauthorized",
      message: error instanceof Error ? error.message : "Token validation failed."
    });
  }
}

function toPrincipal(payload: JWTPayload): AuthenticatedPrincipal {
  return {
    tenantId: String(payload.tid ?? ""),
    objectId: optionalClaim(payload.oid),
    clientId: optionalClaim(payload.azp ?? payload.appid),
    scopes: toClaimList(payload.scp),
    roles: toClaimList(payload.roles)
  };
}

function hasPermission(
  principal: AuthenticatedPrincipal,
  permissionSet: { scopes: string[]; roles: string[] }
) {
  const scopeGranted = permissionSet.scopes.some((scope) => principal.scopes.includes(scope));
  const roleGranted = permissionSet.roles.some((role) => principal.roles.includes(role));
  return scopeGranted || roleGranted;
}

function extractBearerToken(headerValue?: string) {
  if (!headerValue?.startsWith("Bearer ")) {
    return undefined;
  }

  return headerValue.slice("Bearer ".length);
}

function optionalClaim(value: unknown) {
  return typeof value === "string" && value.length > 0 ? value : undefined;
}

function toClaimList(value: unknown) {
  if (typeof value === "string") {
    return value
      .split(" ")
      .map((item) => item.trim())
      .filter(Boolean);
  }

  if (Array.isArray(value)) {
    return value.filter((item): item is string => typeof item === "string" && item.length > 0);
  }

  return [];
}

# Traefik Plugin Bulk Redirects

[![CI](https://github.com/shufanshijie/traefik-bulk-redirects/actions/workflows/pr-build.yaml/badge.svg)](https://github.com/shufanshijie/traefik-bulk-redirects/actions/workflows/pr-build.yaml)
[![release](https://img.shields.io/github/release/shufanshijie/traefik-bulk-redirects/all.svg)](https://github.com/shufanshijie/traefik-bulk-redirects/releases)
[![license](https://img.shields.io/github/license/shufanshijie/traefik-bulk-redirects.svg)](https://github.com/shufanshijie/traefik-bulk-redirects/blob/master/LICENSE)

A Traefik middleware plugin for large, Cloudflare-style redirect sets. It supports exact redirects, query-specific exact redirects, subpath redirects, query string preservation, and configurable redirect status codes.

This repository is a production-oriented fork of [DoodleScheduling/traefik-bulk-redirects](https://github.com/DoodleScheduling/traefik-bulk-redirects). Upstream `v0.0.1` supports inline rules only. This fork is based on the upstream `master` file-mode implementation and adds query-specific matching, duplicate detection, large-rule tests, and a production maintenance workflow.

## Matching behavior

Each request is evaluated in this order:

1. Exact `host + path + raw query` lookup when the request has a query string.
2. Exact `host + path` lookup as a fallback.
3. Most-specific subpath lookup.
4. Pass the request to the next handler.

Exact lookups use Go maps and are average O(1). File rules are loaded, validated, and compiled once, then reused from memory. Requests do not read the rules file.

Matching normalization is intentionally limited and predictable:

- Host names are case-insensitive.
- Explicit ports are ignored.
- The source scheme is not part of the lookup key.
- Paths use their escaped representation.
- Query-specific rules compare the raw query exactly. `a=1&b=2` and `b=2&a=1` are different rules.
- A request that misses a query-specific rule can still match the corresponding rule without a query.
- Source fragments are rejected because browsers do not send fragments to the server.
- Query-specific sources cannot use `subpathMatching`.
- Duplicate normalized keys are rejected at startup instead of silently overwriting a rule.

An exact rule and a subpath rule may intentionally share the same source path. The exact rule handles that path and the subpath rule handles its descendants.

## Redirect fields

| Key | Description |
| --- | --- |
| `sourceURL` | Absolute source URL. Query strings are allowed for exact rules; fragments are rejected. |
| `targetURL` | Absolute redirect destination URL. |
| `statusCode` | Redirect status: `301`, `302`, `303`, `307`, or `308`. Defaults to `301`. |
| `preserveQueryString` | Appends the request's original raw query to the target URL. |
| `subpathMatching` | Matches the source path and child paths. Cannot be combined with a query-specific source. |

## Install

Pin a released version in Traefik's static configuration. Do not deploy from `master` or use a floating tag.

```yaml
experimental:
  plugins:
    companyBulkRedirects:
      moduleName: github.com/shufanshijie/traefik-bulk-redirects
      version: v0.1.0
```

Traefik downloads the plugin during startup. The Git tag, `go.mod` module path, and `.traefik.yml` import path must match exactly.

## Inline mode

Inline mode is the default and is suitable for small rule sets.

```yaml
apiVersion: traefik.io/v1alpha1
kind: Middleware
metadata:
  name: bulk-redirects
spec:
  plugin:
    companyBulkRedirects:
      redirects:
        - sourceURL: https://legacy.example.com/product/100
          targetURL: https://www.example.com/products/100
          statusCode: 301
          preserveQueryString: false
          subpathMatching: false
        - sourceURL: https://legacy.example.com/landing?campaign=spring
          targetURL: https://www.example.com/spring-offer
          statusCode: 301
          preserveQueryString: false
          subpathMatching: false
```

## File mode

File mode is recommended for large rule sets. `filePath` must be absolute and point to a regular JSON file no larger than 16 MiB. Inline `redirects` and `filePath` cannot be combined.

```yaml
apiVersion: traefik.io/v1alpha1
kind: Middleware
metadata:
  name: bulk-redirects
spec:
  plugin:
    companyBulkRedirects:
      mode: file
      filePath: /etc/traefik/bulk-redirects/redirects.json
```

The JSON file accepts only the top-level `redirects` field:

```json
{
  "redirects": [
    {
      "sourceURL": "https://legacy.example.com/product/100",
      "targetURL": "https://www.example.com/products/100",
      "statusCode": 301,
      "preserveQueryString": false,
      "subpathMatching": false
    }
  ]
}
```

The compiled rules are cached by `filePath` for the lifetime of the Traefik process. Replacing a file at the same path does not reload it. Apply every rule update through a Traefik restart or rolling deployment.

## Production rule maintenance

Keep plugin code and redirect data in separate repositories:

- This repository publishes immutable plugin versions such as `v0.1.0`.
- A separate rules repository owns `rules/redirects.csv`, generates `redirects.json`, and publishes an immutable rules image.

Use UTF-8 CSV as the editable source of truth. A practical schema is:

```csv
rule_id,source_url,target_url,status_code,preserve_query_string,subpath_matching,enabled,owner,ticket,updated_at,note
R000001,https://legacy.example.com/product/100,https://www.example.com/products/100,301,false,false,true,seo,SEO-1234,2026-09-28,product migration
```

### Excel to JSON converter

The repository includes a dependency-free Python converter for `.xlsx` and `.xlsm` workbooks. It recognizes both of these Chinese workbook layouts:

- A raw source sheet named `源表` with `页面URL` and `新网站映射链接` columns. When no match-mode column exists, every populated row is treated as an exact redirect.
- A reviewed rules sheet with `规则ID`, `匹配方式`, `匹配串(精确URL或正则)`, and `301目标(新站)` columns.

English columns such as `rule_id`, `source_url`, and `target_url` are also supported.

Run the converter in strict mode first:

```bash
python3 scripts/xlsx_to_redirects.py \
  --input /path/to/seo-mapping.xlsx \
  --sheet 源表 \
  --output dist/redirects.json \
  --min-rules 20000 \
  --max-rules 26000
```

Strict mode rejects unsupported match modes, source fragments, normalized duplicate sources, conflicting targets, invalid URLs, unexpected rule counts, and JSON larger than 16 MiB. It never writes a partial output file.

After reviewing the workbook, the following options can handle specific known cases:

```bash
python3 scripts/xlsx_to_redirects.py \
  --input /path/to/seo-mapping.xlsx \
  --sheet 源表 \
  --output dist/redirects.json \
  --strip-source-fragments \
  --deduplicate-same-target \
  --min-rules 20000 \
  --max-rules 26000
```

- `--skip-unsupported` omits rows such as regular-expression rules and reports the count. The plugin does not support regular expressions.
- `--strip-source-fragments` removes source fragments because browsers do not send them to Traefik.
- `--deduplicate-same-target` keeps the first normalized source only when every duplicate has the same target.
- Conflicting targets for the same normalized `host + path + raw query` always fail and must be resolved in the workbook.

Do not use `--skip-unsupported` for a raw `源表` without a match-mode column. Those rows are already treated as exact redirects. Run strict mode first, review the reported fragments and duplicates, and only then use the two explicit cleanup options above. URLs that differ only by scheme or fragment can still resolve to the same runtime key, so conflicting targets require a business decision before deployment.

The output is deterministic and reports the selected sheet, generated count, skipped count, byte size, and SHA-256. Commit the maintained workbook to the rules repository, but deploy only the generated `redirects.json`.

CI for the rules repository should:

1. Validate required fields, booleans, status codes, absolute URLs, fragments, and query/subpath conflicts.
2. Normalize source keys with the same host, path, and raw-query rules used by the plugin; the scheme and request port are not part of the runtime key.
3. Reject duplicate `rule_id` values and duplicate normalized source keys.
4. Enforce an expected rule-count range to catch accidental bulk deletion.
5. Generate deterministic `redirects.json` and report its rule count, byte size, and SHA-256.
6. Build a rules image tagged by commit SHA and deploy it by immutable image digest.

For approximately 25,000 rules, the JSON commonly exceeds Kubernetes' 1 MiB ConfigMap limit. Store the JSON in a small rules image, copy it to an `emptyDir` with an init container, and mount that directory read-only into Traefik.

```yaml
spec:
  template:
    spec:
      initContainers:
        - name: install-redirect-rules
          image: registry.example.com/infra/traefik-redirect-rules@sha256:<digest>
          command: ["cp", "/rules/redirects.json", "/work/redirects.json"]
          volumeMounts:
            - name: redirect-rules
              mountPath: /work
      containers:
        - name: traefik
          volumeMounts:
            - name: redirect-rules
              mountPath: /etc/traefik/bulk-redirects
              readOnly: true
      volumes:
        - name: redirect-rules
          emptyDir: {}
```

Recommended release flow:

1. Business or SEO staff update the CSV in a pull request.
2. CI validates and generates JSON; a technical owner reviews deletions and target-domain changes.
3. Merge and build the rules image.
4. Update the GitOps deployment to the new image digest.
5. Roll out one canary pod and run known redirect smoke cases.
6. Complete the rolling deployment with `maxUnavailable: 0` and `maxSurge: 1`.

Rollback by restoring the previous rules image digest and rolling Traefik again. Record the plugin version, rules commit, rules image digest, and JSON SHA-256 for each production deployment.

## Development and verification

```bash
go test ./...
go test -race ./...
go test -run '^$' -bench BenchmarkExactRedirectLookup25000 -benchmem ./...
```

The test suite covers query priority and fallback, raw-query ordering, duplicate detection, the 16 MiB file limit, immutable file-cache behavior, concurrent access, loading 25,000 exact rules, and random lookups across a 25,000-rule map.

Before a production rollout, measure total Traefik pod memory with the real rules file. The rule map, JSON decoding peak, Yaegi plugin runtime, Traefik, connections, and observability buffers all contribute to the pod limit.

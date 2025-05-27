# Publishing to npm

## Prerequisites

1. Create an npm account at https://www.npmjs.com/signup
2. Create an organization named "shmaxi" at https://www.npmjs.com/org/create

## Steps to Publish

1. **Login to npm:**
   ```bash
   npm login
   ```
   Enter your username, password, and email when prompted.

2. **Verify you're logged in:**
   ```bash
   npm whoami
   ```

3. **Publish the package:**
   ```bash
   cd /Users/shmax/work/my-excel/excel-mcp-server-node
   npm publish --access public
   ```
   
   Note: The `--access public` flag is required for scoped packages to be public.

## After Publishing

Test the package works with npx:
```bash
npx @shmaxi/excel-mcp-server --help
```

## Updating the Package

1. Update the version in package.json
2. Run `npm publish` again

## Version Management

Use semantic versioning:
- Patch release (bug fixes): `npm version patch`
- Minor release (new features): `npm version minor`
- Major release (breaking changes): `npm version major`

Then publish:
```bash
npm publish
```
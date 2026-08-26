// Production contract template. Back this interface with a tenant/user-scoped
// database or service; do not use a shared filesystem JSON file in production.
export async function get({ tenant, subject }) {
  throw new Error(`Configure a production settings store for ${tenant}/${subject}.`);
}

export async function put({ tenant, subject }, value) {
  void value;
  throw new Error(`Configure a production settings store for ${tenant}/${subject}.`);
}

export async function remove({ tenant, subject }) {
  throw new Error(`Configure a production settings store for ${tenant}/${subject}.`);
}

export async function deleteSettings(key) {
  return remove(key);
}

export default { get, put, delete: deleteSettings };

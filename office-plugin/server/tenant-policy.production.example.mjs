// Copy this file to a separately managed production module and replace the
// placeholder implementation with an admin-backed tenant policy service.
export async function allows({ identity, capability }) {
  if (!identity?.tenant || !identity?.subject) return false;
  // Fail closed until the tenant policy service is connected.
  return false;
}

export default { allows };

// Production contract template. Replace with a shared Redis/service-backed
// limiter before enabling more than one application instance.
export async function consume({ tenant, subject, route }) {
  void tenant;
  void subject;
  void route;
  throw new Error("Configure a distributed production rate limiter.");
}

export default { consume };

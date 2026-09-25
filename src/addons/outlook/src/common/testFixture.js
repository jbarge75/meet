// TEST ONLY — used on this test branch to bypass ProConnect auth and real
// room creation, so the Polycom URL-param toggle can be tested manually
// without a backend round-trip. Do not merge into the feature branch.
const { BASE_URL } = require("./index");

const TEST_ROOM_DATA = {
  id: "rif-pymx-ndj",
  name: "rif-pymx-ndj",
  slug: "rif-pymx-ndj",
  url: `${BASE_URL}/rif-pymx-ndj`,
  telephony: {
    enabled: true,
    phone_number: null,
    pin_code: "0425282403",
    default_country: "FR",
  },
};

module.exports = { TEST_ROOM_DATA };

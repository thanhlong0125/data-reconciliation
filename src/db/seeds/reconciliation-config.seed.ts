import type { ReconciliationConfigDocument } from '../models/reconciliation-config.model';

export const RECONCILIATION_CONFIG_SEED: ReconciliationConfigDocument[] = [
  {
    code: 1,
    name: 'App - Thẻ quốc tế (MPGS)',
    partner: {
      key: '',
      price: '',
      time: '',
      refund: '',
      cutKeyPartner: {
        key: 'abc',
        index: 1,
      },
    },
    afc: {
      key: '',
      price: '',
      time: '',
    },
    active: true,
  },
];

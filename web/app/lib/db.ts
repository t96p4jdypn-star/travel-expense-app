import Dexie, { type EntityTable } from "dexie";
import { normalizeState, type AppState } from "./types";

type StoredState = { id: "current"; data: AppState; updatedAt: string };

class TravelExpenseDatabase extends Dexie {
  states!: EntityTable<StoredState, "id">;
  constructor() {
    super("travel-expense-app");
    this.version(1).stores({ states: "id, updatedAt" });
  }
}

export const db = new TravelExpenseDatabase();

export async function loadState(): Promise<AppState | null> {
  const data = (await db.states.get("current"))?.data;
  return data ? normalizeState(data) : null;
}

export async function saveState(data: AppState): Promise<void> {
  await db.states.put({ id: "current", data, updatedAt: new Date().toISOString() });
}

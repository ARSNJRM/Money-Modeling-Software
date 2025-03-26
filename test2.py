from dataclasses import dataclass, field
from typing import Dict, List, Optional
import json, openpyxl
import pandas as pd

# Define a class to represent an entry in an agent's balance sheet
@dataclass
class BalanceSheetEntry:
    entry_type: str  # "asset" or "liability"
    category: str    # e.g., "loan", "bond"
    value_per_unit: float
    quantity: int
    counterparty: Optional[str]
    due_date: str    # "tX" 或 "on-demand"
    means_of_settlement: Optional[str]

# Define a class representing an agent (e.g., banks, companies) in the system
@dataclass
class Agent:
    id: str
    balance_sheet: Dict[str, List[BalanceSheetEntry]] = field(default_factory=dict)
    status: str = "operating"

    # Method to add a new financial entry to the agent's balance sheet
    def add_entry(self, entry: BalanceSheetEntry):
        if entry.entry_type not in self.balance_sheet:
            self.balance_sheet[entry.entry_type] = []
        self.balance_sheet[entry.entry_type].append(entry)

    # Retrieve all assets of the agent
    def get_assets(self):
        return self.balance_sheet.get("asset", [])
    
    # Retrieve all liabilities of the agent
    def get_liabilities(self):
        return self.balance_sheet.get("liability", [])
    
    # Calculate the agent's cash balance
    def get_cash_balance(self) -> float:
        return sum(
            entry.value_per_unit * entry.quantity 
            for entry in self.get_assets()
            if entry.category == "cash" and entry.due_date == "on-demand"
        )

# Define the main financial system for managing agents, transactions, and central bank operations
class MoneyModelingSystem:
    def __init__(self):
        self.agents: Dict[str, Agent] = {}
        self.transaction_history: Dict[str, List[str]] = {}
        self._init_central_bank()

    def _get_means_of_settlement(self) -> Optional[str]:
        while True:
            choice = input("Specify means of settlement? (specified/unspecified): ").lower()
            if choice == "specified":
                asset_type = input("Which asset? (e.g., cash, bond, machines): ")
                return asset_type
            elif choice == "unspecified":
                return None
            else:
                print("Invalid choice. Please enter 'specified' or 'unspecified'.")

    def _record_transaction(self, timepoint: str, description: str):
        if timepoint not in self.transaction_history:
            self.transaction_history[timepoint] = []
        self.transaction_history[timepoint].append(description)

    def add_agent(self, agent_id: str):
        if agent_id not in self.agents:
            self.agents[agent_id] = Agent(agent_id)
            return self.agents[agent_id]
        return None
    
    # Initialize the central bank with a default cash liability
    def _init_central_bank(self):
        cb = Agent("CentralBank")
        cb.add_entry(BalanceSheetEntry(
            entry_type="liability",
            category="cash",
            value_per_unit=1,
            quantity=0,
            counterparty=None,
            due_date="on-demand",
            means_of_settlement=None
        ))
        self.agents["CentralBank"] = cb

    def issue_cash(self, receiver_id: str, amount: float):
        if receiver_id not in self.agents:
            raise ValueError(f"Agent {receiver_id} not found")
        
        self.agents[receiver_id].add_entry(BalanceSheetEntry(
            entry_type="asset",
            category="cash",
            value_per_unit=1.0, 
            quantity=amount,  
            counterparty="CentralBank",
            due_date="on-demand",
            means_of_settlement=None
        ))

        cb = self.agents["CentralBank"]
        cash_liability = next(
            entry for entry in cb.get_liabilities()
            if entry.category == "cash"
        )
        cash_liability.quantity += amount  

    # Create a new transaction between agents
    def create_transaction(self):
        while True:
            direction = input("Create from (asset/liability): ").lower()
            if direction in ["asset", "liability"]:
                break
            print("Invalid choice. Please enter 'asset' or 'liability'")

        category = input("Transaction type (e.g., loan, bond): ") # change to asset/liability type
        value = float(input("Value per unit: "))
        quantity = int(input("Quantity: "))
        due_date = input("Due date (tX format or on-demand): ")
        means = self._get_means_of_settlement()

        if direction == "asset":
            holder = input("Asset holder (owner) ID: ")
            issuer = input("Liability issuer (counterparty) ID: ")
            self._create_asset_transaction(
                category=category,
                value=value,
                quantity=quantity,
                holder=holder,
                issuer=issuer,
                due_date=due_date,
                means=means
            )
        else:
            issuer = input("Liability issuer ID: ")
            holder = input("Asset holder (counterparty) ID: ")
            self._create_liability_transaction(
                category=category,
                value=value,
                quantity=quantity,
                issuer=issuer,
                holder=holder,
                due_date=due_date,
                means=means
            )

    def _create_asset_transaction(self, category: str, value: float, quantity: int,
                                holder: str, issuer: str, due_date: str, means: Optional[str]):
        self.agents[holder].add_entry(BalanceSheetEntry(
            entry_type="asset",
            category=category,
            value_per_unit=value,
            quantity=quantity,
            counterparty=issuer,
            due_date=due_date,
            means_of_settlement=means
        ))
        
        self.agents[issuer].add_entry(BalanceSheetEntry(
            entry_type="liability",
            category=category,
            value_per_unit=value,
            quantity=quantity,
            counterparty=holder,
            due_date=due_date,
            means_of_settlement=means
        ))

        total_value = value * quantity
        self._record_transaction("t0",
            f"{holder} created {category} asset ({issuer}) {total_value} due {due_date}")
        self._record_transaction("t0",
            f"{issuer} created {category} liability ({holder}) {total_value} due {due_date}")

    def _create_liability_transaction(self, category: str, value: float, quantity: int,
                                    issuer: str, holder: str, due_date: str, means: Optional[str]):
        self.agents[issuer].add_entry(BalanceSheetEntry(
            entry_type="liability",
            category=category,
            value_per_unit=value,
            quantity=quantity,
            counterparty=holder,
            due_date=due_date,
            means_of_settlement=means
        ))
        
        self.agents[holder].add_entry(BalanceSheetEntry(
            entry_type="asset",
            category=category,
            value_per_unit=value,
            quantity=quantity,
            counterparty=issuer,
            due_date=due_date,
            means_of_settlement=means
        ))

        total_value = value * quantity
        self._record_transaction("t0",
            f"{issuer} created {category} liability ({holder}) {total_value} due {due_date}")
        self._record_transaction("t0",
            f"{holder} created {category} asset ({issuer}) {total_value} due {due_date}")
    
    # Process and settle transactions at a given timepoint
    def process_timepoint(self, timepoint: str):
        print(f"\nProcessing timepoint {timepoint}")
        for agent in list(self.agents.values()):
            if agent.status != "operating":
                continue

            if agent.id == "CentralBank":
                continue

            for liability in agent.get_liabilities():
                if liability.due_date == timepoint:
                    if liability.means_of_settlement is None:
                        print(f"Unspecified settlement: {agent.id} can use any asset")
                        self._settle_with_any_asset(agent, liability)
                    else:
                        print(f"Specified settlement: {agent.id} must use {liability.means_of_settlement}")
                        self._settle_with_specific_asset(agent, liability)
    
    def _settle_with_specific_asset(self, agent: Agent, liability: BalanceSheetEntry):
        required = liability.value_per_unit * liability.quantity
        available = sum(
            entry.value_per_unit * entry.quantity
            for entry in agent.get_assets()
            if entry.category == liability.means_of_settlement
        )
        
        if available >= required:
            print(f"Settled using {liability.means_of_settlement}")
            timepoint = liability.due_date 
            self._transfer_cash(
                payer=agent.id,
                payee=liability.counterparty,
                asset_type=liability.means_of_settlement,
                amount=required,
                timepoint=timepoint
            )
        else:
            print("Default occurred!")
            agent.status = "bankrupt"

    def _settle_with_any_asset(self, agent: Agent, liability: BalanceSheetEntry):
        required = liability.value_per_unit * liability.quantity
        total_assets = sum(
            entry.value_per_unit * entry.quantity
            for entry in agent.get_assets()
        )
        
        if total_assets >= required:
            print("Settled using any available assets")
        else:
            print("Default occurred!")
            agent.status = "bankrupt"

    def _transfer_cash(self, payer: str, payee: str, asset_type: str , amount: float, timepoint: str):
        payer_agent = self.agents[payer]
        cash_entries = [
            entry for entry in payer_agent.get_assets()
            if entry.category == "cash" 
            and entry.due_date == "on-demand"
        ]
        
        remaining = amount
        for entry in cash_entries:
            total_value = entry.value_per_unit * entry.quantity
            if total_value >= remaining:
                entry.quantity -= int(remaining / entry.value_per_unit)
                remaining = 0
                break
            else:
                remaining -= total_value
                entry.quantity = 0

        self.agents[payee].add_entry(BalanceSheetEntry(
            entry_type="asset",
            category="cash",
            value_per_unit=1.0,
            quantity=amount,
            counterparty="CentralBank",
            due_date="on-demand",
            means_of_settlement=None
        ))
        self._record_transaction(timepoint,
            f"{payer} -> {payee}: -Cash ${amount}")
        self._record_transaction(timepoint,
            f"{payee} -> {payer}: +Cash ${amount}")

    def _settle_payment(self, payer_id: str, liability: BalanceSheetEntry):
        self.agents[payer_id].balance_sheet["liability"].remove(liability)
        
        receiver = self.agents[liability.counterparty]
        matching_asset = next(
            a for a in receiver.get_assets()
            if a.category == liability.category
            and a.amount == liability.amount
            and a.counterparty == payer_id
        )
        receiver.balance_sheet["asset"].remove(matching_asset)

    def _remove_entry(self, agent: Agent, entry: BalanceSheetEntry):
        agent.balance_sheet[entry.entry_type].remove(entry)

    # Print transaction history in a table format
    def print_transaction_table(self):
        print("\nTransaction History:")
        for tp in sorted(self.transaction_history.keys()):
            print(f"\nTimepoint {tp}:")
            print(f"{'Agent':<15} | {'Assets':<40} | {'Liabilities':<40}")
            print("-"*100)
            
            agent_states = {}
            for desc in self.transaction_history[tp]:
                if "created" in desc:
                    parts = desc.split()
                    agent = parts[0]
                    entry_type = "asset" if "asset" in desc else "liability"
                    details = " ".join(parts[2:-3])
                    amount = parts[-3]
                    due = parts[-1]
                    
                    if agent not in agent_states:
                        agent_states[agent] = {"assets": [], "liabilities": []}
                        
                    if entry_type == "asset":
                        agent_states[agent]["assets"].append(f"+{details} {amount} ({due})")
                    else:
                        agent_states[agent]["liabilities"].append(f"+{details} {amount} ({due})")
                
                elif "->" in desc:
                    agents, amount = desc.split(":")
                    payer, payee = agents.split(" -> ")
                    amount = amount.strip()
                    
                    for agent in [payer, payee]:
                        if agent not in agent_states:
                            agent_states[agent] = {"assets": [], "liabilities": []}
                    
                    agent_states[payer]["assets"].append(amount)
                    agent_states[payee]["assets"].append(amount.replace("-", "+"))

            for agent, entries in agent_states.items():
                assets = "\n".join(entries["assets"])
                liabilities = "\n".join(entries["liabilities"])
                print(f"{agent:<15} | {assets:<40} | {liabilities:<40}")

    def export_transaction_history_to_excel(self, filename: str = "transaction_history.xlsx"):
        data = []
        for timepoint, transactions in self.transaction_history.items():
            row = {"Timepoint": timepoint}
            for transaction in transactions:
                parts = transaction.split()
                agent = parts[0]
                entry_type = parts[2]  # "asset" or "liability"
                details = " ".join(parts[3:]) 
                
                if agent not in row:
                    row[agent] = {}
                row[agent][entry_type] = details
            
            data.append(row)
        
        print(data)

        df = pd.DataFrame(data)
        df = df.set_index("Timepoint").apply(pd.Series.explode).reset_index()
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Transactions")
            
            worksheet = writer.sheets["Transactions"]
            
            for column in worksheet.columns:
                max_length = max(len(str(cell.value)) for cell in column)
                adjusted_width = max_length + 2 
                worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
            
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.alignment = openpyxl.styles.Alignment(wrap_text=True)
        
        print(f"Transaction history exported to {filename}")


def initialize_system(system: MoneyModelingSystem):
    print("""
    Money Modeling Software v1.0 (with Central Bank)
    ------------------------------------------------
    1. Create agents
    2. Add transactions
    3. Issue cash
    4. Run simulation
    5. Show analysis
    6. Export to Excel
    7. Exit
    """)

    while True:
        choice = input("Select operation (1-7): ")
        
        if choice == "1":
            agent_id = input("Enter agent ID: ")
            if system.add_agent(agent_id):
                print(f"Agent {agent_id} created")
            else:
                print("Agent already exists!")
        
        elif choice == "2":
            system.create_transaction() 
            print("Transaction created")
        
        elif choice == "3": 
            receiver = input("Receiver agent ID: ")
            amount = float(input("Amount to issue: "))
            try:
                system.issue_cash(receiver, amount)
                print(f"Successfully issued {amount} cash to {receiver}")
            except Exception as e:
                print(f"Error: {str(e)}")
        
        elif choice == "4":
            timepoints = sorted({entry.due_date for agent in system.agents.values() 
                               for entry in agent.get_liabilities()})
            for tp in timepoints:
                if tp == "on-demand":
                    continue
                if not system.process_timepoint(tp):
                    print("System halted due to default")
                    break
            else:
                print("Simulation completed successfully")
        
        elif choice == "5":
            print("\nSystem Analysis:")
            system.print_transaction_table()
            print(f"Central Bank Cash Liability: {system.agents['CentralBank'].get_liabilities()[0].value_per_unit * system.agents['CentralBank'].get_liabilities()[0].quantity}")
            for agent in system.agents.values():
                print(f"{agent.id}:")
                print(f"  Cash Balance: {agent.get_cash_balance()}")
                print(f"  Status: {agent.status}")
        
        elif choice == "6":
            filename = input("Enter output filename (e.g., transactions.xlsx): ")
            system.export_transaction_history_to_excel(filename)
        
        elif choice == "7":
            break


if __name__ == "__main__":
    system = MoneyModelingSystem()
    
    # # 创建示例代理
    # system.add_agent("Bank1")
    # system.add_agent("Company1")
    
    # # 中央银行发行现金
    # system.issue_cash("Bank1", 1000)
    # system.issue_cash("Company1", 500)
    
    # # 创建贷款交易
    # system.create_transaction(
    #     entry_type="asset",
    #     category="loan",
    #     amount=800,
    #     from_agent="Bank1",
    #     to_agent="Company1",
    #     due_date="t2",
    #     means_of_payment="cash"
    # )
    
    initialize_system(system)
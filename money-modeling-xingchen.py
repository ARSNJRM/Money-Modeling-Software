"""
Economic System Simulation Module 

This module provides classes and functions for simulating an economic system with various agents,
assets, liabilities, and financial instruments.
"""

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple, Union
from enum import Enum
from datetime import datetime, timedelta
from copy import deepcopy

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    EXCEL_AVAILABLE = True
except ImportError:
    print("Warning: openpyxl package not found. Excel export functionality will be disabled.")
    print("To enable Excel export, please install openpyxl using: pip install openpyxl")
    EXCEL_AVAILABLE = False

# Core Enums and Classes
class AgentType(Enum):
    """Types of economic agents"""
    BANK = "bank"
    COMPANY = "company"
    HOUSEHOLD = "household"
    TREASURY = "treasury"
    CENTRAL_BANK = "central_bank"
    OTHER = "other"

class EntryType(Enum):
    LOAN = "loan"
    DEPOSIT = "deposit"
    PAYABLE = "payable"
    BOND_ZERO_COUPON = "bond_zero_coupon"
    BOND_COUPON = "bond_coupon"
    BOND_AMORTIZING = "bond_amortizing"
    DELIVERY_CLAIM = "delivery_claim"
    NON_FINANCIAL = "non_financial"
    DEFAULT = "default"

class MaturityType(Enum):
    """Types of maturity for financial instruments"""
    ON_DEMAND = "on_demand"
    FIXED_DATE = "fixed_date"
    PERPETUAL = "perpetual"

class SettlementType(Enum):
    """Types of settlement for financial instruments"""
    MEANS_OF_PAYMENT = "means_of_payment"
    SECURITIES = "securities"
    NON_FINANCIAL_ASSET = "non_financial_asset"
    SERVICES = "services"
    CRYPTO = "crypto"
    NONE = "none"

class BondType(Enum):
    """Types of bonds"""
    ZERO_COUPON = 0  # Zero-coupon bond
    COUPON = 1      # Regular coupon bond
    AMORTIZING = 2  # Amortizing bond

@dataclass
class SettlementDetails:
    type: SettlementType
    denomination: str

@dataclass
class BalanceSheetEntry:
    type: EntryType
    is_asset: bool
    counterparty: Optional[str]
    amount: float
    denomination: str
    maturity_type: MaturityType
    maturity_date: Optional[datetime]
    settlement_details: SettlementDetails
    name: Optional[str] = None
    issuance_time: str = 't0'
    book_value: Optional[float] = None
    expected_cash_flow: Optional[float] = None
    parent_bond: Optional[str] = None  # Reference to the main bond

    def matches(self, other: 'BalanceSheetEntry') -> bool:
        return (
            self.type == other.type and
            self.is_asset == other.is_asset and
            self.counterparty == other.counterparty and
            self.amount == other.amount and
            self.denomination == other.denomination and
            self.maturity_type == other.maturity_type and
            self.maturity_date == other.maturity_date and
            self.settlement_details.type == other.settlement_details.type and
            self.settlement_details.denomination == other.settlement_details.denomination and
            self.name == other.name and
            self.issuance_time == other.issuance_time
        )

    def __post_init__(self):
        if self.amount <= 0:
            raise ValueError("Amount must be positive")
        if self.issuance_time not in ['t0', 't1', 't2']:
            raise ValueError("Issuance time must be 't0', 't1', or 't2'")
        if self.type != EntryType.NON_FINANCIAL and not self.counterparty:
            raise ValueError("Counterparty is required for financial entries")
        if self.type == EntryType.NON_FINANCIAL and self.counterparty:
            raise ValueError("Non-financial entries cannot have a counterparty")
        if self.type == EntryType.NON_FINANCIAL and not self.name:
            raise ValueError("Non-financial entries must have a name")
        if self.type == EntryType.PAYABLE and self.settlement_details.type != SettlementType.MEANS_OF_PAYMENT:
            raise ValueError("Payable entries must have means_of_payment settlement type")

class Agent:
    """Represents an economic agent in the system"""
    def __init__(self, name: str, agent_type: AgentType):
        self.name = name
        self.type = agent_type
        self.assets: List[BalanceSheetEntry] = []
        self.liabilities: List[BalanceSheetEntry] = []
        self.status: str = "operating"
        self.creation_time: datetime = datetime.now()
        self.settlement_history = {
            'as_asset_holder': [],
            'as_liability_holder': []
        }

    def add_asset(self, entry: BalanceSheetEntry):
        self.assets.append(entry)

    def add_liability(self, entry: BalanceSheetEntry):
        self.liabilities.append(entry)

    def remove_asset(self, entry: BalanceSheetEntry):
        self.assets = [e for e in self.assets if not e.matches(entry)]

    def remove_liability(self, entry: BalanceSheetEntry):
        self.liabilities = [e for e in self.liabilities if not e.matches(entry)]

    def get_total_assets(self) -> float:
        return sum(entry.amount for entry in self.assets)

    def get_total_liabilities(self) -> float:
        return sum(entry.amount for entry in self.liabilities)

    def get_net_worth(self) -> float:
        return self.get_total_assets() - self.get_total_liabilities()

    def record_settlement(self, time_point: str, original_entry: BalanceSheetEntry,
                         settlement_result: BalanceSheetEntry, counterparty: str,
                         as_asset_holder: bool):
        settlement_record = {
            'time_point': time_point,
            'original_entry': deepcopy(original_entry),
            'settlement_result': deepcopy(settlement_result),
            'counterparty': counterparty,
            'timestamp': datetime.now()
        }
        if as_asset_holder:
            self.settlement_history['as_asset_holder'].append(settlement_record)
        else:
            self.settlement_history['as_liability_holder'].append(settlement_record)

class AssetLiabilityPair:
    """Represents a pair of corresponding asset and liability entries"""
    def __init__(self, time: datetime, type: str, amount: float,
                 denomination: str, maturity_type: MaturityType,
                 maturity_date: Optional[datetime], settlement_type: SettlementType,
                 settlement_denomination: str, asset_holder: Agent,
                 liability_holder: Optional[Agent] = None,
                 asset_name: Optional[str] = None,
                 bond_type: Optional[int] = None,
                 coupon_rate: Optional[float] = None):
        self.time = time
        self.type = type
        self.amount = amount
        self.denomination = denomination
        self.maturity_type = maturity_type
        self.maturity_date = maturity_date
        self.settlement_details = SettlementDetails(
            type=settlement_type,
            denomination=settlement_denomination
        )
        self.coupon_rate = coupon_rate
        self.bond_type = bond_type
        self.asset_holder = asset_holder
        self.liability_holder = liability_holder
        self.asset_name = asset_name
        self.initial_book_value = amount  # BV₀
        self.connected_claims = []  # 存储相关的claims

        if type == EntryType.NON_FINANCIAL.value:
            if liability_holder is not None:
                raise ValueError("Non-financial entries cannot have a liability holder")
            if not asset_name:
                raise ValueError("Non-financial entries must have an asset name")
        else:
            if liability_holder is None:
                raise ValueError("Financial entries must have a liability holder")

    def create_entries(self) -> Tuple[BalanceSheetEntry, Optional[BalanceSheetEntry]]:
        """Create corresponding asset and liability entries"""
        # Verify that only banks can hold loans
        if self.type == EntryType.LOAN.value:
            if self.asset_holder.type != AgentType.BANK:
                raise ValueError("Only banks can hold loans as assets")
        
        # Create basic balance sheet entries
        asset_entry = BalanceSheetEntry(
            type=EntryType(self.type),
            is_asset=True,
            counterparty=self.liability_holder.name if self.liability_holder else None,
            amount=self.amount,
            denomination=self.denomination,
            maturity_type=self.maturity_type,
            maturity_date=self.maturity_date,
            settlement_details=self.settlement_details,
            name=self.asset_name
        )

        if self.type == EntryType.NON_FINANCIAL.value:
            return asset_entry, None

        liability_entry = BalanceSheetEntry(
            type=EntryType(self.type),
            is_asset=False,
            counterparty=self.asset_holder.name,
            amount=self.amount,
            denomination=self.denomination,
            maturity_type=self.maturity_type,
            maturity_date=self.maturity_date,
            settlement_details=self.settlement_details,
            name=None
        )

        # Create connected claims for coupon and amortizing bonds
        if self.type in [EntryType.BOND_COUPON.value, EntryType.BOND_AMORTIZING.value]:
            claims = self.create_bond_claims()
            self.connected_claims = claims
            
            # Add claims to asset holder's assets
            for claim in claims:
                self.asset_holder.add_asset(claim)
                
                # Create corresponding liability
                liability = BalanceSheetEntry(
                    type=claim.type,
                    is_asset=False,
                    counterparty=self.asset_holder.name,
                    amount=claim.amount,
                    denomination=claim.denomination,
                    maturity_type=claim.maturity_type,
                    maturity_date=claim.maturity_date,
                    settlement_details=claim.settlement_details
                )
                self.liability_holder.add_liability(liability)

        return asset_entry, liability_entry

    def _create_bond_payment_schedule(self) -> List[Tuple[datetime, float, str]]:
        """Create payment schedule for different types of bonds"""
        if not self.bond_type:
            raise ValueError("Bond type must be specified for bonds")

        schedule = []
        t1 = datetime(2050, 1, 1)
        t2 = datetime(2100, 1, 1)
        
        if self.type == EntryType.BOND_ZERO_COUPON.value:
            # Zero-coupon bond: only principal payment at maturity
            schedule.append((self.maturity_date, self.amount, "principal"))

        elif self.type == EntryType.BOND_COUPON.value:
            if not self.coupon_rate:
                raise ValueError("Coupon rate is required for coupon bonds")
            
            is_t2_maturity = self.maturity_date == t2
            
            if not is_t2_maturity:  # Matures at t1
                schedule.append((t1, self.amount * self.coupon_rate, 'coupon'))
                schedule.append((t1, self.amount, 'principal'))
            else:  # Matures at t2
                schedule.append((t1, self.amount * self.coupon_rate, 'coupon'))
                schedule.append((t2, self.amount * self.coupon_rate, 'coupon'))
                schedule.append((t2, self.amount, 'principal'))

        elif self.type == EntryType.BOND_AMORTIZING.value:
            if not self.coupon_rate:
                raise ValueError("Interest rate is required for amortizing bonds")
            
            is_t2_maturity = self.maturity_date == t2
            
            if is_t2_maturity:  # Matures at t2
                # Split payments between t1 and t2
                principal_payment_t1 = self.amount / 2
                interest_payment_t1 = self.amount * self.coupon_rate
                schedule.append((t1, principal_payment_t1, 'principal'))
                schedule.append((t1, interest_payment_t1, 'coupon'))
                
                principal_payment_t2 = self.amount / 2
                interest_payment_t2 = (self.amount / 2) * self.coupon_rate
                schedule.append((t2, principal_payment_t2, 'principal'))
                schedule.append((t2, interest_payment_t2, 'coupon'))
            else:  # Matures at t1
                schedule.append((t1, self.amount, 'principal'))
                schedule.append((t1, self.amount * self.coupon_rate, 'coupon'))

        # Debug information
        print(f"\nBond Payment Schedule Details:")
        print(f"Bond Type: {self.type}")
        print(f"Maturity Date: {'t2' if self.maturity_date == t2 else 't1'}")
        print("\nPayment Schedule:")
        for date, amount, payment_type in sorted(schedule):
            time_point = 't2' if date == t2 else 't1'
            print(f"- {time_point}: {amount:.2f} ({payment_type})")

        return schedule

    def _create_bond_entries(self, payment_schedule: List[Tuple[datetime, float, str]]) -> Tuple[BalanceSheetEntry, BalanceSheetEntry]:
        """Create bond entries"""
        # Create main bond entries
        asset_entry = BalanceSheetEntry(
            type=EntryType(self.type),
            is_asset=True,
            counterparty=self.liability_holder.name,
            amount=self.amount,
            denomination=self.denomination,
            maturity_type=self.maturity_type,
            maturity_date=self.maturity_date,
            settlement_details=self.settlement_details,
            name=f"{self.type} bond"
        )

        liability_entry = BalanceSheetEntry(
            type=EntryType(self.type),
            is_asset=False,
            counterparty=self.asset_holder.name,
            amount=self.amount,
            denomination=self.denomination,
            maturity_type=self.maturity_type,
            maturity_date=self.maturity_date,
            settlement_details=self.settlement_details
        )

        return asset_entry, liability_entry

    def _calculate_amortization_payment(self) -> float:
        """Calculate the payment amount for amortizing bond"""
        r = self.coupon_rate  # annual interest rate
        pv = self.amount      # face value
        
        # Since our model only has one payment at t1 or t2
        # We return the sum of principal and interest due
        total_payment = pv * (1 + r)
        
        return total_payment

    def create_bond_claims(self) -> List[BalanceSheetEntry]:
        """Create connected claims for bonds"""
        claims = []
        if self.type in [EntryType.BOND_COUPON.value, EntryType.BOND_AMORTIZING.value]:
            schedule = self._create_bond_payment_schedule()
            
            # Generate a unique identifier for this bond
            bond_id = f"bond_{self.type}_{id(self)}"  # 使用对象的id作为唯一标识符
            
            for date, amount, payment_type in schedule:
                # Calculate book value and expected cash flow
                bv = amount  # This should be calculated based on the formula
                cf = amount / self.initial_book_value  # As portion of BV₀
                
                # Create claim with correct maturity date
                claim = BalanceSheetEntry(
                    type=EntryType.PAYABLE,
                    is_asset=True,
                    counterparty=self.liability_holder.name,
                    amount=amount,
                    denomination=self.denomination,
                    maturity_type=MaturityType.FIXED_DATE,
                    maturity_date=date,
                    settlement_details=self.settlement_details,
                    name=f"{payment_type}_claim_{date.year}",  # 不再依赖bond的name属性
                    book_value=bv,
                    expected_cash_flow=cf,
                    parent_bond=bond_id  # 使用生成的bond_id
                )
                claims.append(claim)
                
                print(f"Creating claim for {payment_type} payment at {date.year}")
                
        return claims

    def transfer_to(self, new_holder: Agent):
        """Transfer bond and all connected claims to new holder"""
        # Transfer main bond entry
        self.asset_holder.remove_asset(self.asset_entry)
        new_holder.add_asset(self.asset_entry)
        
        # Transfer all connected claims
        for claim in self.connected_claims:
            self.asset_holder.remove_asset(claim)
            new_holder.add_asset(claim)
        
        self.asset_holder = new_holder

class EconomicSystem:
    """Main class for managing the economic system simulation"""
    def __init__(self):
        self.agents: Dict[str, Agent] = {}
        self.asset_liability_pairs: List[AssetLiabilityPair] = []
        self.time_states: Dict[str, Dict[str, Agent]] = {}
        self.current_time_state = "t0"
        self.simulation_finalized = False
        self.save_state('t0')
        self.system_dates = {
            't1': datetime(2050, 1, 1),  # 这个可以在系统初始化时设置
            't2': datetime(2100, 1, 1)
        }

    def add_agent(self, agent: Agent):
        self.agents[agent.name] = agent
        if self.current_time_state == 't0':
            self.save_state('t0')

    def create_asset_liability_pair(self, pair: AssetLiabilityPair):
        # Verify that only banks can hold loans (this check is already in create_entries)
        asset_entry, liability_entry = pair.create_entries()
        self.asset_liability_pairs.append(pair)
        pair.asset_holder.add_asset(asset_entry)
        if liability_entry:
            pair.liability_holder.add_liability(liability_entry)

        # When creating a loan, automatically create corresponding deposit
        if pair.type == EntryType.LOAN.value:
            # At this point, we're guaranteed that asset_holder is a bank
            # because create_entries() would have thrown an error otherwise
            deposit_pair = AssetLiabilityPair(
                time=datetime.now(),
                type=EntryType.DEPOSIT.value,
                amount=pair.amount,  # Deposit amount equals loan amount
                denomination=pair.denomination,
                maturity_type=MaturityType.ON_DEMAND,
                maturity_date=None,
                settlement_type=SettlementType.MEANS_OF_PAYMENT,
                settlement_denomination=pair.denomination,
                asset_holder=pair.liability_holder,  # Borrower gets deposit asset
                liability_holder=pair.asset_holder,  # Bank holds deposit liability
            )
            
            deposit_asset, deposit_liability = deposit_pair.create_entries()
            deposit_pair.asset_holder.add_asset(deposit_asset)
            deposit_pair.liability_holder.add_liability(deposit_liability)
            self.asset_liability_pairs.append(deposit_pair)

        self.save_state(self.current_time_state)

    def save_state(self, time_point: str):
        if time_point not in ['t0', 't1', 't2']:
            raise ValueError("Invalid time point")
        self.time_states[time_point] = {}
        for name, agent in self.agents.items():
            agent_copy = Agent(agent.name, agent.type)
            agent_copy.assets = deepcopy(agent.assets)
            agent_copy.liabilities = deepcopy(agent.liabilities)
            self.time_states[time_point][name] = agent_copy
        self.current_time_state = time_point

    def display_balance_sheets(self, time_point: str):
        if not self.agents:
            print("\nNo agents in the system!")
            return

        current_agents = self.get_agents_at_time(time_point).values()
        print(f"\nBalance sheet at {time_point}:")
        print("Assets:")
        for agent in current_agents:
            for asset in agent.assets:
                maturity_info = ""
                if asset.maturity_type == MaturityType.FIXED_DATE:
                    if asset.maturity_date.year == 2100:
                        maturity_info = " (matures at t2)"
                    elif asset.maturity_date.year == 2050:
                        maturity_info = " (matures at t1)"

                entry_type = asset.type.value
                if asset.type == EntryType.PAYABLE:
                    entry_type = "Receivable"
                elif asset.type == EntryType.DELIVERY_CLAIM:
                    entry_type = f"Delivery claim {asset.name}" if asset.name else "Delivery claim"
                elif asset.type == EntryType.DEFAULT:
                    entry_type = f"Default claim ({asset.name})"

                print(f"  - {entry_type}: {asset.amount} {asset.denomination} "
                      f"(from {asset.counterparty if asset.counterparty else 'N/A'})"
                      f"{maturity_info} [issued at {asset.issuance_time}]")

            print("Liabilities:")
            for liability in agent.liabilities:
                maturity_info = ""
                if liability.maturity_type == MaturityType.FIXED_DATE:
                    if liability.maturity_date.year == 2100:
                        maturity_info = " (matures at t2)"
                    elif liability.maturity_date.year == 2050:
                        maturity_info = " (matures at t1)"

                entry_type = liability.type.value
                if liability.type == EntryType.DELIVERY_CLAIM:
                    entry_type = f"Delivery promise {liability.name}" if liability.name else "Delivery promise"
                elif liability.type == EntryType.DEFAULT:
                    entry_type = f"Default liability ({liability.name})"

                print(f"  - {entry_type}: {liability.amount} {liability.denomination} "
                      f"(to {liability.counterparty}){maturity_info} "
                      f"[issued at {liability.issuance_time}]")

            print(f"Total assets: {agent.get_total_assets()}")
            print(f"Total liabilities: {agent.get_total_liabilities()}")
            print(f"Net worth: {agent.get_net_worth()}")

    def get_agents_at_time(self, time_point: str) -> Dict[str, Agent]:
        if time_point not in ['t0', 't1', 't2']:
            raise ValueError("Invalid time point")
        if time_point == 't0':
            return {name: agent for name, agent in self.agents.items()}
        if time_point in self.time_states:
            return self.time_states[time_point]
        return {name: agent for name, agent in self.agents.items()}

    def run_simulation(self) -> bool:
        print("\nStarting simulation...")
        self.save_state('t0')
        print("t0 state saved")
        return True

    def settle_entries(self, time_point: str):
        """Settle entries at the specified time point"""
        self.validate_time_point(time_point, allow_t0=False)

        # First save the state of the previous time point
        prev_time = 't0' if time_point == 't1' else 't1'
        if prev_time not in self.time_states:
            self.save_state(prev_time)

        # Process all matured entries
        for pair in self.asset_liability_pairs[:]:  # Create a copy for iteration
            if pair.maturity_type == MaturityType.FIXED_DATE and pair.maturity_date:
                # Check if the entry's maturity date matches the current time state
                entry_time = 't1' if pair.maturity_date.year == 2050 else 't2'
                
                if entry_time == time_point:
                    # Handle bond payments
                    if pair.type.startswith("bond_"):
                        self._settle_bond(pair, time_point)
                    else:
                        # Handle other types of settlement
                        self._settle_non_bond(pair, time_point)

        # Automatically save state after settlement
        self.save_state(time_point)
        self.current_time_state = time_point

    def _settle_bond(self, pair: AssetLiabilityPair, time_point: str):
        """Handle bond settlement"""
        # Get payment schedule
        payment_schedule = pair._create_bond_payment_schedule()
        
        # Find amounts to be paid at current time point
        current_payments = [
            (amount, payment_type) 
            for date, amount, payment_type in payment_schedule 
            if (time_point == 't1' and date.year == 2050) or 
               (time_point == 't2' and date.year == 2100)
        ]

        # Process each payment
        for amount, payment_type in current_payments:
            # Check if debtor has sufficient deposit
            debtor_deposit = next(
                (asset for asset in pair.liability_holder.assets
                 if asset.type == EntryType.DEPOSIT
                 and asset.amount >= amount),
                None
            )

            if not debtor_deposit:
                raise ValueError(f"Debtor does not have enough deposit to pay {payment_type}")

            # Get bank holding the deposit liability
            bank = next(a for a in self.agents.values() if a.name == debtor_deposit.counterparty)

            # Deduct payment amount from debtor's deposit
            pair.liability_holder.remove_asset(debtor_deposit)
            bank.remove_liability(next(
                l for l in bank.liabilities
                if l.counterparty == pair.liability_holder.name
                and l.amount == debtor_deposit.amount
            ))

            # Create new deposit for debtor (if there's remaining amount)
            if debtor_deposit.amount > amount:
                remainder_deposit = AssetLiabilityPair(
                    time=datetime.now(),
                    type=EntryType.DEPOSIT.value,
                    amount=debtor_deposit.amount - amount,
                    denomination=pair.denomination,
                    maturity_type=MaturityType.ON_DEMAND,
                    maturity_date=None,
                    settlement_type=SettlementType.NONE,
                    settlement_denomination=pair.denomination,
                    asset_holder=pair.liability_holder,
                    liability_holder=bank
                )
                remainder_asset, remainder_liability = remainder_deposit.create_entries()
                remainder_asset.issuance_time = time_point
                remainder_liability.issuance_time = time_point
                pair.liability_holder.add_asset(remainder_asset)
                bank.add_liability(remainder_liability)

            # Create new deposit for creditor
            creditor_deposit = AssetLiabilityPair(
                time=datetime.now(),
                type=EntryType.DEPOSIT.value,
                amount=amount,
                denomination=pair.denomination,
                maturity_type=MaturityType.ON_DEMAND,
                maturity_date=None,
                settlement_type=SettlementType.NONE,
                settlement_denomination=pair.denomination,
                asset_holder=pair.asset_holder,
                liability_holder=bank
            )
            creditor_asset, creditor_liability = creditor_deposit.create_entries()
            creditor_asset.issuance_time = time_point
            creditor_liability.issuance_time = time_point
            pair.asset_holder.add_asset(creditor_asset)
            bank.add_liability(creditor_liability)

            # Record settlement history
            pair.asset_holder.record_settlement(
                time_point=time_point,
                original_entry=pair.create_entries()[0],  # Original bond asset
                settlement_result=creditor_asset,  # New deposit
                counterparty=pair.liability_holder.name,
                as_asset_holder=True
            )
            pair.liability_holder.record_settlement(
                time_point=time_point,
                original_entry=pair.create_entries()[1],  # Original bond liability
                settlement_result=remainder_asset if debtor_deposit.amount > amount else None,
                counterparty=pair.asset_holder.name,
                as_asset_holder=False
            )

        # If this is the last payment (maturity), remove original bond
        if time_point == 't2' or (time_point == 't1' and pair.maturity_date.year == 2050):
            # Remove original bond entries
            asset_entry, liability_entry = pair.create_entries()
            pair.asset_holder.remove_asset(asset_entry)
            if liability_entry:
                pair.liability_holder.remove_liability(liability_entry)
            
            # Remove bond pair from system
            if pair in self.asset_liability_pairs:
                self.asset_liability_pairs.remove(pair)

    def _settle_non_bond(self, pair: AssetLiabilityPair, time_point: str):
        """Handle non-bond type settlement"""
        # Remove original entries
        asset_entry, liability_entry = pair.create_entries()
        pair.asset_holder.remove_asset(asset_entry)
        if liability_entry:
            pair.liability_holder.remove_liability(liability_entry)

        # Handle settlement
        if pair.settlement_details.type == SettlementType.MEANS_OF_PAYMENT:
            # Find deposit for settlement
            debtor_deposit = next(
                (asset for asset in pair.liability_holder.assets
                 if asset.type == EntryType.DEPOSIT
                 and asset.amount >= pair.amount
                 and asset.denomination == pair.denomination),
                None
            )

            if not debtor_deposit:
                raise ValueError(f"No suitable deposit found for settlement")

            # Get bank holding the deposit liability
            bank = next(a for a in self.agents.values() if a.name == debtor_deposit.counterparty)

            # Remove original deposit from debtor
            pair.liability_holder.remove_asset(debtor_deposit)

            # Remove corresponding liability from bank
            bank_liability = next(
                (l for l in bank.liabilities
                 if l.type == EntryType.DEPOSIT
                 and l.counterparty == pair.liability_holder.name
                 and l.amount == debtor_deposit.amount),
                None
            )
            if bank_liability:
                bank.remove_liability(bank_liability)

            # Create new deposit entry for creditor
            settlement_pair = AssetLiabilityPair(
                time=datetime.now(),
                type=EntryType.DEPOSIT.value,
                amount=pair.amount,
                denomination=pair.denomination,
                maturity_type=MaturityType.ON_DEMAND,
                maturity_date=None,
                settlement_type=SettlementType.NONE,
                settlement_denomination=pair.denomination,
                asset_holder=pair.asset_holder,  # Creditor gets deposit
                liability_holder=bank,  # Bank maintains liability
                asset_name=None
            )

            # Create entries with current time point as issuance time
            new_asset_entry, new_liability_entry = settlement_pair.create_entries()
            new_asset_entry.issuance_time = time_point
            if new_liability_entry:
                new_liability_entry.issuance_time = time_point

            # Record settlement history
            pair.asset_holder.record_settlement(
                time_point=time_point,
                original_entry=asset_entry,
                settlement_result=new_asset_entry,
                counterparty=pair.liability_holder.name,
                as_asset_holder=True
            )
            pair.liability_holder.record_settlement(
                time_point=time_point,
                original_entry=liability_entry,
                settlement_result=debtor_deposit,  # Used deposit
                counterparty=pair.asset_holder.name,
                as_asset_holder=False
            )

            # Add entries
            settlement_pair.asset_holder.add_asset(new_asset_entry)
            if new_liability_entry:
                settlement_pair.liability_holder.add_liability(new_liability_entry)
            self.asset_liability_pairs.append(settlement_pair)

            # If there's remaining deposit amount, create new deposit for remainder
            if debtor_deposit.amount > pair.amount:
                remainder_pair = AssetLiabilityPair(
                    time=datetime.now(),
                    type=EntryType.DEPOSIT.value,
                    amount=debtor_deposit.amount - pair.amount,
                    denomination=pair.denomination,
                    maturity_type=MaturityType.ON_DEMAND,
                    maturity_date=None,
                    settlement_type=SettlementType.NONE,
                    settlement_denomination=pair.denomination,
                    asset_holder=pair.liability_holder,  # Debtor keeps remainder
                    liability_holder=bank,  # Bank maintains liability
                    asset_name=None
                )

                # Create entries with current time point as issuance time
                remainder_asset, remainder_liability = remainder_pair.create_entries()
                remainder_asset.issuance_time = time_point
                if remainder_liability:
                    remainder_liability.issuance_time = time_point

                # Add entries
                remainder_pair.asset_holder.add_asset(remainder_asset)
                if remainder_liability:
                    remainder_pair.liability_holder.add_liability(remainder_liability)
                self.asset_liability_pairs.append(remainder_pair)

        elif pair.settlement_details.type == SettlementType.NON_FINANCIAL_ASSET:
            # Find and remove non-financial asset from liability holder
            non_financial_asset = next(
                (asset for asset in pair.liability_holder.assets
                 if asset.type == EntryType.NON_FINANCIAL
                 and asset.name == pair.asset_name
                 and asset.amount >= pair.amount),
                None
            )

            if not non_financial_asset:
                raise ValueError(f"Non-financial asset {pair.asset_name} not found for settlement")

            # Remove asset from liability holder
            pair.liability_holder.remove_asset(non_financial_asset)

            # Create non-financial asset entry for asset holder
            settlement_pair = AssetLiabilityPair(
                time=datetime.now(),
                type=EntryType.NON_FINANCIAL.value,
                amount=pair.amount,
                denomination=pair.settlement_details.denomination,
                maturity_type=MaturityType.ON_DEMAND,
                maturity_date=None,
                settlement_type=SettlementType.NONE,
                settlement_denomination=pair.settlement_details.denomination,
                asset_holder=pair.asset_holder,  # Original creditor gets goods
                liability_holder=None,  # Non-financial assets have no liability holder
                asset_name=pair.asset_name  # Use asset name directly
            )

            # Create entry with current time point as issuance time
            new_asset_entry, _ = settlement_pair.create_entries()
            new_asset_entry.issuance_time = time_point

            # Record settlement history
            pair.asset_holder.record_settlement(
                time_point=time_point,
                original_entry=asset_entry,
                settlement_result=new_asset_entry,
                counterparty=pair.liability_holder.name,
                as_asset_holder=True
            )
            pair.liability_holder.record_settlement(
                time_point=time_point,
                original_entry=liability_entry,
                settlement_result=non_financial_asset,  # Delivered non-financial asset
                counterparty=pair.asset_holder.name,
                as_asset_holder=False
            )

            # Add entry directly to avoid default t0 issuance time
            settlement_pair.asset_holder.add_asset(new_asset_entry)
            self.asset_liability_pairs.append(settlement_pair)

            # If non-financial asset has remaining amount, create new entry for it
            if non_financial_asset.amount > pair.amount:
                remainder_pair = AssetLiabilityPair(
                    time=datetime.now(),
                    type=EntryType.NON_FINANCIAL.value,
                    amount=non_financial_asset.amount - pair.amount,
                    denomination=non_financial_asset.denomination,
                    maturity_type=MaturityType.ON_DEMAND,
                    maturity_date=None,
                    settlement_type=SettlementType.NONE,
                    settlement_denomination=non_financial_asset.denomination,
                    asset_holder=pair.liability_holder,  # Original holder keeps remainder
                    liability_holder=None,
                    asset_name=non_financial_asset.name
                )

                # Create entry with current time point as issuance time
                remainder_asset, _ = remainder_pair.create_entries()
                remainder_asset.issuance_time = time_point

                # Add entry
                remainder_pair.asset_holder.add_asset(remainder_asset)
                self.asset_liability_pairs.append(remainder_pair)

        # Remove original pair from system
        if pair in self.asset_liability_pairs:
            self.asset_liability_pairs.remove(pair)

    def validate_time_point(self, time_point: str, allow_t0: bool = True) -> None:
        """Validate time point string"""
        valid_points = ['t0', 't1', 't2'] if allow_t0 else ['t1', 't2']
        if time_point not in valid_points:
            raise ValueError(f"Time point must be {', '.join(valid_points)}")

    def create_asset_liability_pair_interactive(self):
        """
        Interactive function to create an asset-liability pair.
        Handles user input for all necessary parameters and validates bank requirement for loans.
        """
        # Get asset holder first
        print("\nAvailable asset holders:")
        for i, agent in enumerate(self.agents.values(), 1):
            print(f"{i}. {agent.name} ({agent.type.value})")
        asset_holder_idx = int(input("Select asset holder (enter number): ")) - 1
        asset_holder = list(self.agents.values())[asset_holder_idx]

        # Get liability holder
        print("\nAvailable liability holders:")
        for i, agent in enumerate(self.agents.values(), 1):
            if agent != asset_holder:  # Cannot select self as liability holder
                print(f"{i}. {agent.name} ({agent.type.value})")
        liability_holder_idx = int(input("Select liability holder (enter number): ")) - 1
        liability_holder = list(self.agents.values())[liability_holder_idx]

        # Get entry type
        print("\nAvailable entry types:")
        print("1. Loan")
        print("2. Deposit")
        print("3. Receivable-payable")
        print("4. Bond")
        print("5. Delivery claim")
        
        type_idx = int(input("Select entry type (enter number): ")) - 1

        # Handle different entry types
        if type_idx == 0:  # loan
            if asset_holder.type != AgentType.BANK:
                print("Error: Only banks can hold loans as assets!")
                return
            entry_type = EntryType.LOAN
            bond_type = None
            coupon_rate = None
            
        elif type_idx == 1:  # deposit
            entry_type = EntryType.DEPOSIT
            bond_type = None
            coupon_rate = None
            
        elif type_idx == 2:  # receivable-payable
            entry_type = EntryType.PAYABLE
            bond_type = None
            coupon_rate = None
            
        elif type_idx == 3:  # bond
            print("\nSelect bond type:")
            print("1. Zero-coupon bond (No periodic interest payments)")
            print("2. Coupon bond (Periodic interest payments)")
            print("3. Amortizing bond (Gradual principal repayment)")
            
            while True:
                try:
                    bond_type_idx = int(input("Enter choice (1-3): ")) - 1
                    if bond_type_idx not in [0, 1, 2]:
                        print("Error: Please enter a number between 1 and 3")
                        continue
                    break
                except ValueError:
                    print("Error: Please enter a valid number")
            
            if bond_type_idx == 0:
                entry_type = EntryType.BOND_ZERO_COUPON
                bond_type = BondType.ZERO_COUPON
                coupon_rate = None
                print("\nNote: Zero-coupon bond will be issued at a discount and redeemed at face value")
            elif bond_type_idx == 1:
                entry_type = EntryType.BOND_COUPON
                bond_type = BondType.COUPON
                while True:
                    try:
                        coupon_rate = float(input("\nEnter annual coupon rate (as decimal, e.g. 0.05 for 5%): "))
                        if coupon_rate <= 0:
                            print("Error: Coupon rate must be positive")
                            continue
                        print(f"\nConfirmed: Bond will pay {coupon_rate*100}% interest periodically")
                        break
                    except ValueError:
                        print("Error: Please enter a valid number")
            else:
                entry_type = EntryType.BOND_AMORTIZING
                bond_type = BondType.AMORTIZING
                while True:
                    try:
                        coupon_rate = float(input("\nEnter interest rate (as decimal, e.g. 0.05 for 5%): "))
                        if coupon_rate <= 0:
                            print("Error: Interest rate must be positive")
                            continue
                        print(f"\nConfirmed: Principal will be gradually repaid with {coupon_rate*100}% interest")
                        break
                    except ValueError:
                        print("Error: Please enter a valid number")
        elif type_idx == 4:  # delivery claim
            entry_type = EntryType.DELIVERY_CLAIM
            bond_type = None
            coupon_rate = None
        else:
            print("Invalid choice!")
            return

        # Get amount and denomination
        while True:
            try:
                amount = float(input("\nEnter amount (face value): "))
                if amount <= 0:
                    print("Error: Amount must be positive")
                    continue
                break
            except ValueError:
                print("Error: Please enter a valid number")
                
        denomination = input("Enter denomination (e.g., USD, EUR): ")

        # For bonds, we require fixed maturity date
        if entry_type in [EntryType.BOND_ZERO_COUPON, EntryType.BOND_COUPON, EntryType.BOND_AMORTIZING]:
            maturity_type = MaturityType.FIXED_DATE
            print("\nSelect bond maturity:")
            print("1. t1 (2050)")
            print("2. t2 (2100)")
            time_point = int(input("Enter choice (1-2): "))
            maturity_date = datetime(2050, 1, 1) if time_point == 1 else datetime(2100, 1, 1)
        else:
            # Get maturity type and date for non-bond entries
            print("\nSelect maturity type:")
            print("1. On demand")
            print("2. Fixed date")
            print("3. Perpetual")
            maturity_idx = int(input("Enter choice (1-3): ")) - 1
            
            if maturity_idx == 0:
                maturity_type = MaturityType.ON_DEMAND
                maturity_date = None
            elif maturity_idx == 1:
                maturity_type = MaturityType.FIXED_DATE
                print("\nSelect maturity time point:")
                print("1. t1 (2050)")
                print("2. t2 (2100)")
                time_point = int(input("Enter choice (1-2): "))
                maturity_date = datetime(2050, 1, 1) if time_point == 1 else datetime(2100, 1, 1)
            else:
                maturity_type = MaturityType.PERPETUAL
                maturity_date = None

        # For bonds, default to means of payment settlement
        if entry_type in [EntryType.BOND_ZERO_COUPON, EntryType.BOND_COUPON, EntryType.BOND_AMORTIZING]:
            settlement_type = SettlementType.MEANS_OF_PAYMENT
            settlement_denomination = denomination
        else:
            # Get settlement type and denomination for non-bond entries
            print("\nSelect settlement type:")
            print("1. Means of payment")
            print("2. Securities")
            print("3. Non-financial asset")
            print("4. Services")
            print("5. Crypto")
            print("6. None")
            settlement_idx = int(input("Enter choice (1-6): ")) - 1
            settlement_type = list(SettlementType)[settlement_idx]
            settlement_denomination = input("Enter settlement denomination (e.g., USD, EUR): ")

        try:
            # Create asset-liability pair
            pair = AssetLiabilityPair(
                time=datetime.now(),
                type=entry_type.value,
                amount=amount,
                denomination=denomination,
                maturity_type=maturity_type,
                maturity_date=maturity_date,
                settlement_type=settlement_type,
                settlement_denomination=settlement_denomination,
                asset_holder=asset_holder,
                liability_holder=liability_holder,
                asset_name=None,
                bond_type=bond_type,
                coupon_rate=coupon_rate
            )

            self.create_asset_liability_pair(pair)
            if entry_type in [EntryType.BOND_ZERO_COUPON, EntryType.BOND_COUPON, EntryType.BOND_AMORTIZING]:
                print("\nBond created successfully!")
                # Display payment schedule
                schedule = pair._create_bond_payment_schedule()
                print("\nPayment schedule:")
                for date, amount, payment_type in schedule:
                    print(f"- {date.strftime('%Y')}: {amount} {denomination} ({payment_type})")
            else:
                print("\nAsset-liability pair created successfully!")
            
        except ValueError as e:
            print(f"\nError creating asset-liability pair: {str(e)}")
            return

if __name__ == "__main__":
    def main():
        # Initialize the economic system
        system = EconomicSystem()

        while True:
            print("\nAvailable actions:")
            print("1. Add new agent")
            print("2. Create asset-liability pair")
            print("3. Display balance sheets")
            print("4. Run simulation to t1")
            print("5. Run simulation to t2")
            print("6. Exit")

            try:
                choice = int(input("\nSelect action (enter number): "))
                
                if choice == 1:
                    print("\nSelect agent type:")
                    print("1. Bank")
                    print("2. Company")
                    print("3. Household")
                    print("4. Treasury")
                    print("5. Central Bank")
                    print("6. Other")
                    
                    agent_type_idx = int(input("Enter choice (1-6): ")) - 1
                    agent_name = input("Enter agent name: ")
                    
                    agent_type = list(AgentType)[agent_type_idx]
                    new_agent = Agent(agent_name, agent_type)
                    system.add_agent(new_agent)
                    print(f"\nAgent {agent_name} ({agent_type.value}) added successfully!")

                elif choice == 2:
                    if len(system.agents) < 2:
                        print("\nError: Need at least 2 agents to create an asset-liability pair!")
                        continue
                    system.create_asset_liability_pair_interactive()

                elif choice == 3:
                    print("\nSelect time point:")
                    print("1. t0")
                    print("2. t1")
                    print("3. t2")
                    time_idx = int(input("Enter choice (1-3): "))
                    time_point = ['t0', 't1', 't2'][time_idx - 1]
                    system.display_balance_sheets(time_point)

                elif choice == 4:
                    if system.run_simulation():
                        system.settle_entries('t1')
                        print("Simulation to t1 completed")

                elif choice == 5:
                    if 't1' not in system.time_states:
                        print("\nError: Must run simulation to t1 first!")
                        continue
                    system.settle_entries('t2')
                    print("Simulation to t2 completed")

                elif choice == 6:
                    print("\nExiting program...")
                    break

                else:
                    print("\nInvalid choice!")

            except ValueError as e:
                print(f"\nError: {str(e)}")
            except Exception as e:
                print(f"\nUnexpected error: {str(e)}")

    main() 

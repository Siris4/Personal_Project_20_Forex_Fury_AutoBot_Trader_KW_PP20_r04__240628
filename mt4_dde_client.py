import win32com.client
import logging

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def set_stop_loss(dde, symbol, open_price, ticket, stop_loss_percentage=2):
    """Calculate and set stop loss for an order."""
    stop_loss_price = open_price * (1 - stop_loss_percentage / 100)
    try:
        dde.Poke(f"{symbol}.SL.{ticket}", stop_loss_price)
        logger.info(f"Set stop loss for order {ticket} at {stop_loss_price}")
    except Exception as e:
        logger.error(f"Error setting stop loss for order {ticket}: {e}")

def main():
    try:
        # Create a DDE client
        dde = win32com.client.Dispatch("DDEClient.DDEClient")
        logger.info("DDE client created.")
    except Exception as e:
        logger.error(f"Error initializing DDE client: {e}")
        return

    try:
        # Connect to MT4's DDE server
        dde.Connect("MT4", "BID")
        logger.info("Connected to MT4 DDE server.")
    except Exception as e:
        logger.error(f"Error connecting to MT4 DDE server: {e}")
        return

    try:
        # Request data
        eurusd_bid = dde.Request("EURUSD")
        logger.info("Data requested from MT4 DDE server.")
        print("EURUSD Bid: ", eurusd_bid)
    except Exception as e:
        logger.error(f"Error requesting data from MT4 DDE server: {e}")

    try:
        # Example to set stop loss (replace with actual trade details)
        # Fetch open trades and set stop loss for each
        trades = [
            {'Symbol': 'EURUSD', 'OpenPrice': 1.1234, 'Ticket': '12345'},
            # Add more trades as needed
        ]
        for trade in trades:
            set_stop_loss(dde, trade['Symbol'], trade['OpenPrice'], trade['Ticket'])
    except Exception as e:
        logger.error(f"Error setting stop loss for trades: {e}")

    try:
        # Disconnect
        dde.Disconnect()
        logger.info("Disconnected from MT4 DDE server.")
    except Exception as e:
        logger.error(f"Error disconnecting from MT4 DDE server: {e}")

if __name__ == "__main__":
    main()

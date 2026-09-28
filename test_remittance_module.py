from app import app, db, User, RemittanceRecord
from datetime import date

def test_endpoints():
    with app.test_client() as client:
        # 1. Login as Admin
        with client.session_transaction() as sess:
            admin_user = User.query.filter_by(role='Admin').first()
            if not admin_user:
                print("No admin user found")
                return
            sess['user_id'] = admin_user.id
            sess['_fresh'] = True

        print(f"Logged in as user {admin_user.username} (ID: {admin_user.id})")

        # 2. Test Remittance Dashboard
        resp = client.get('/remittances')
        assert resp.status_code == 200, f"/remittances returned {resp.status_code}"
        assert b"Remittance, Cash Incentive" in resp.data, "Dashboard title missing"
        assert b"020122" in resp.data, "Invoice 020122 missing"
        print("[PASS] GET /remittances -> 200 OK (Dashboard loads with sample records)")

        # 3. Test Create Page GET
        resp = client.get('/remittances/create')
        assert resp.status_code == 200, f"/remittances/create returned {resp.status_code}"
        assert b"Record Remittance, Incentive" in resp.data
        print("[PASS] GET /remittances/create -> 200 OK")

        # 4. Test Create Entry POST
        post_data = {
            'date': '2022-06-01',
            'particulars': 'Remittance Received',
            'invoice_no': 'TEST-INV-99',
            'sent_by': 'Global Tech Client',
            'deal_no': 'DL-8822',
            'remittance_usd': '15000.00',
            'usd_rate': '86.50',
            'received_bdt': '1297500.00',
            'incentive_10_bdt': '129750.00',
            'incentive_10_audited_bdt': '129750.00',
            'tt_certificate': 'on',
            'encashment_certificate': 'on',
            'status': 'Under Bank Review'
        }
        resp = client.post('/remittances/create', data=post_data, follow_redirects=True)
        assert resp.status_code == 200
        assert b"TEST-INV-99" in resp.data
        print("[PASS] POST /remittances/create -> 200 OK (Entry created & visible on dashboard)")

        # Clean up the test entry
        test_rec = RemittanceRecord.query.filter_by(invoice_no='TEST-INV-99').first()
        if test_rec:
            rec_id = test_rec.id
            # 5. Test Edit Page GET
            resp = client.get(f'/remittances/{rec_id}/edit')
            assert resp.status_code == 200
            assert b"Edit Entry" in resp.data
            print("[PASS] GET /remittances/<id>/edit -> 200 OK")

            # 6. Test Checkpoint Toggle API
            resp = client.post('/remittances/api/toggle-checkpoint', data={'id': rec_id, 'field': 'send_to_bank', 'val': 'true'})
            assert resp.status_code == 200
            print("[PASS] POST /remittances/api/toggle-checkpoint -> 200 OK")

            # Delete the test record
            db.session.delete(test_rec)
            db.session.commit()
            print("[PASS] Cleaned up test record")

        # 7. Test Statement View
        resp = client.get('/remittances/statement')
        assert resp.status_code == 200, f"/remittances/statement returned {resp.status_code}"
        assert b"Statement of Inward Remittance" in resp.data
        print("[PASS] GET /remittances/statement -> 200 OK (Printable statement rendered)")

        # 8. Test Corporate Budget Dashboard
        resp = client.get('/budgets')
        assert resp.status_code == 200
        assert b"TBN" in resp.data or b"Corporate Multi-Division" in resp.data
        print("[PASS] GET /budgets -> 200 OK (Corporate budget format rendered)")

        print("\nALL TESTS PASSED SUCCESSFULLY!")

if __name__ == '__main__':
    with app.app_context():
        test_endpoints()

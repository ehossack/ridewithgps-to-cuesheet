from decimal import Decimal, InvalidOperation

import pytest

from ridewithgps_to_cuesheet.conversion import Cue, EventDetails, GenerationOptions, _map_direction, _parse_to_cues


def test_default_options():
    opts = GenerationOptions()

    assert opts.include_distance_from_last is False
    assert opts.hide_direction is False
    assert opts.verbose is False
    assert "Control" in opts.control_cue_indicators
    assert "Start" in opts.control_cue_indicators
    assert "End" in opts.control_cue_indicators
    assert "Summit" in opts.control_cue_indicators
    assert opts.end_indicator == "Summit"
    assert opts.start_text == "DÉPART"
    assert opts.end_text == "ARRIVÉE"
    assert opts.page_break_row_interval == 40


def test_custom_options():
    custom_event = EventDetails(name="Custom Event", date="Custom Date", organizer="Custom Organizer")
    opts = GenerationOptions(
        include_distance_from_last=True,
        hide_direction=True,
        verbose=True,
        page_break_row_interval=20,
        event_details=custom_event,
    )

    assert opts.include_distance_from_last is True
    assert opts.hide_direction is True
    assert opts.verbose is True
    assert opts.page_break_row_interval == 20
    assert opts.event_details.name == "Custom Event"
    # Test default values for non-constructor fields
    assert "Control" in opts.control_cue_indicators
    assert opts.end_indicator == "Summit"


def test_cue_creation():
    cue = Cue(turn="R", description="Turn right onto Main St", dist=Decimal("5.0"), is_control=False)

    assert cue.turn == "R"
    assert cue.description == "Turn right onto Main St"
    assert cue.dist == Decimal("5.0")
    assert cue.is_control is False


def test_cue_with_control():
    cue = Cue(turn="L", description="Food stop at cafe", dist=Decimal("10.0"), is_control=True)

    assert cue.turn == "L"
    assert cue.description == "Food stop at cafe"
    assert cue.dist == Decimal("10.0")
    assert cue.is_control is True


def test_map_basic_turns():
    assert _map_direction("right") == "R"
    assert _map_direction("left") == "L"
    assert _map_direction("straight") == "CO"


def test_map_bear_turns():
    assert _map_direction("slight right") == "BR"
    assert _map_direction("slight left") == "BL"


def test_map_special_turns():
    assert _map_direction("uturn") == "TA"


def test_map_special_cases():
    assert _map_direction("generic") == ""
    assert _map_direction("food") == ""
    assert _map_direction("control") == ""
    assert _map_direction("start") == ""
    assert _map_direction("end") == ""
    assert _map_direction("summit") == ""


def test_map_unknown_turn():
    result = _map_direction("Unknown direction")

    assert result == "Unknown direction"  # Should return the original direction for unknown cases


def test_parse_simple_route():
    csv_data = [
        ["Start", "Start of route", "0", "0", ""],
        ["Right", "Right on Test St", "0.5", "10.0", ""],
        ["Left", "Left on Main St", "2.0", "15.0", ""],
        ["End", "End of route", "5.0", "20.0", ""],
    ]
    opts = GenerationOptions()

    cues = _parse_to_cues(csv_data, opts)

    assert len(cues) == 4
    assert cues[0].turn == ""
    assert cues[0].description == "DÉPART"
    assert cues[0].dist == Decimal("0.5")  # Should be distance from next row (parsing goes backward)
    assert cues[1].turn == "R"
    assert cues[1].description == "Right on Test St"
    assert cues[1].dist == Decimal("2.0")  # Should be distance from next row
    assert cues[-1].turn == ""
    # "End" is in control_cue_indicators but not the end_indicator (which is "Summit")
    assert cues[-1].description == "ARRIVÉE"  # Just the basic end text, not modified


def test_parse_with_controls():
    csv_data = [
        ["Start", "Start of route", "0", "0", ""],
        ["Food", "Food stop at cafe", "5.0", "20.0", ""],
        ["Control", "Control checkpoint", "10.0", "25.0", ""],
        ["End", "End of route", "15.0", "30.0", ""],
    ]
    opts = GenerationOptions()

    cues = _parse_to_cues(csv_data, opts)

    assert len(cues) == 4
    assert cues[1].is_control is False  # Food is not in control_cue_indicators anymore
    assert cues[1].description == "Food stop at cafe"
    assert cues[2].is_control is True  # Control is in control_cue_indicators
    assert cues[2].description == "Control checkpoint"


def test_parse_with_summit_end():
    csv_data = [
        ["Start", "Start of route", "0", "0", ""],
        ["Summit", "Summit at viewpoint", "15.0", "100.0", ""],
        ["End", "End of route", "20.0", "30.0", ""],
    ]
    opts = GenerationOptions()

    cues = _parse_to_cues(csv_data, opts)

    assert len(cues) == 3  # All rows should be processed, not filtered
    assert cues[0].turn == ""
    assert cues[0].description == "DÉPART"
    assert cues[1].turn == ""
    assert cues[1].description == "ARRIVÉE: Summit at viewpoint"  # Summit is the end_indicator


def test_parse_with_custom_end_indicator():
    csv_data = [
        ["Start", "Start of route", "0", "0", ""],
        ["Finish", "Finish line", "10.0", "50.0", ""],
        ["End", "End of route", "15.0", "30.0", ""],
    ]
    opts = GenerationOptions(end_indicator="Finish")

    cues = _parse_to_cues(csv_data, opts)

    assert len(cues) == 3  # All rows should be processed
    assert cues[-1].description != "ARRIVÉE: End of route"  # End is NOT the end_indicator


def test_parse_route_with_turns():
    csv_data = [
        ["Start", "Start of route", "0", "0", ""],
        ["Right", "Turn right onto Test St", "0.5", "10.0", ""],
        ["Left", "Turn left onto Main St", "2.0", "15.0", ""],
        ["Food", "Food stop at cafe", "5.0", "20.0", ""],
        ["Control", "Control checkpoint", "10.0", "25.0", ""],
        ["Right", "At roundabout, take exit 2 onto Sesame Street", "11.0", "25.5", ""],
        ["Slight Left", "Turn slight left onto Big Bird Lane", "12.0", "30.5", ""],
        ["Right", "Keep right onto Big Bird Lane", "12.5", "32.2", ""],
        ["Uturn", "Make a U-turn on Big Bird Lane", "13.5", "35.0", ""],
        ["Left", "Turn left into Bletchley Park", "14.0", "45.0", ""],
        ["Control", "Control 2: Bletchley Park visitor centre", "14.2", "45.2", ""],
        ["End", "End of route", "20.0", "30.0", ""],
    ]
    expected_cues = [
        Cue('', 'DÉPART', Decimal('0.5'), is_control=True, last_dist=Decimal('0')),
        Cue('R', 'Test St', Decimal('2.0'), last_dist=Decimal('0.5')),
        Cue('L', 'Main St', Decimal('5.0'), last_dist=Decimal('2.0')),
        Cue('', 'Food stop at cafe', Decimal('10.0'), last_dist=Decimal('5.0')),
        Cue('', 'Control checkpoint', Decimal('11.0'), is_control=True, last_dist=Decimal('10.0')),
        Cue('R', 'Sesame Street (roundabout exit 2)', Decimal('12.0'), last_dist=Decimal('11.0')),
        Cue('BL', 'Big Bird Lane', Decimal('12.5'), last_dist=Decimal('12.0')),
        Cue('R', 'Big Bird Lane', Decimal('13.5'), last_dist=Decimal('12.5')),
        Cue('TA', 'Big Bird Lane', Decimal('14.0'), last_dist=Decimal('13.5')),
        Cue('L', 'Bletchley Park', Decimal('14.2'), last_dist=Decimal('14.0')),
        Cue('', 'Bletchley Park visitor centre', Decimal('20.0'), is_control=True, last_dist=Decimal('14.2')),
        Cue('', 'ARRIVÉE', Decimal('-1.0'), is_control=True, last_dist=Decimal('20.0')),
    ]
    opts = GenerationOptions()
    cues = _parse_to_cues(csv_data, opts)
    assert cues == expected_cues


def test_parse_empty_data():
    csv_data = []
    opts = GenerationOptions()

    cues = _parse_to_cues(csv_data, opts)

    assert len(cues) == 0


@pytest.mark.parametrize(
    "dist,expected",
    [
        ("invalid", InvalidOperation),
        ("-1.B", InvalidOperation),
    ],
)
def test_parse_invalid_distance(dist, expected):
    csv_data = [
        ["Start", "Start of route", dist, "0", ""],
    ]
    opts = GenerationOptions()

    with pytest.raises(expected):
        _parse_to_cues(csv_data, opts)

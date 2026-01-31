/**
 * @typedef {Object.<string, boolean>} EventDict
 */

export var Constants = (function () {
  function Constants() {}

  // Default values
  Constants.DEFAULT_TOURNAMENT_NAME = "Tournament Name";
  Constants.DEFAULT_TOURNAMENT_DATE = "1-January-2000";
  Constants.DEFAULT_TOURNAMENT_LOCATION = "School_Name";
  Constants.DEFAULT_TOURNAMENT_DIVISION = "__";
  Constants.DEFAULT_NUMBER_OF_TEAMS = "0";
  Constants.DEFAULT_NUMBER_OF_EVENT_MEDALS = 4;
  Constants.DEFAULT_NUMBER_OF_TEAM_TROPIES = 8;

  // Defined named ranges
  Constants.TOURNAMENT_NAME_RANGE_NAME = "TournamentName";
  Constants.TOURNAMENT_DATE_RANGE_NAME = "TournamentDate";
  Constants.LOCATION_RANGE_NAME = "Location";
  Constants.DIVISION_RANGE_NAME = "Division";
  Constants.NUMBER_OF_TEAMS_RANGE_NAME = "NumTeams";
  Constants.NUMBER_OF_EVENT_MEDALS = "NumEventMedals";
  Constants.NUMBER_OF_TEAM_TROPHIES = "NumeTeamTrophies";

  // Last updated for year 2024-25
  /** @type {EventDict} */
  Constants.B_DIVISION_EVENTS = {
    "Air Trajectory": true,
    "Anatomy & Physiology": true,
    "Codebusters": true,
    "Crime Busters": true,
    "Disease Detectives": true,
    "Dynamic Planet": true,
    "Ecology": true,
    "Entomology": true,
    "Experimental Design": true,
    "Fossils": true,
    "Helicopter": true,
    "Meteorology": true,
    "Metric Mastery": true,
    "Microbe Mission": true,
    "Mission Possible": true,
    "Optics": true,
    "Potions & Poisons": true,
    "Reach for the Stars": true,
    "Road Scholar": true,
    "Scrambler": true,
    "Sustainable Energy": true,
    "Tower": true,
    "Write It Do It": true,
    "Fermi Questions": true,
  };

  /** @type {EventDict} */
  Constants.C_DIVISION_EVENTS = {
    "Air Trajectory": true,
    "Anatomy & Physiology": true,
    "Astronomy": true,
    "Chemistry Lab": true,
    "Codebusters": true,
    "Disease Detectives": true,
    "Dynamic Planet": true,
    "Ecology": true,
    "Electric Vehicles": true,
    "Entomology": true,
    "Experimental Design": true,
    "Forensics": true,
    "Fossils": true,
    "Geologic Mapping": true,
    "Helicopter": true,
    "Materials Science": true,
    "Microbe Mission": true,
    "Optics": true,
    "Quantum Quandary": true,
    "Robot Tour": true,
    "Sustainable Energy": true,
    "Tower": true,
    "Write It Do It": true,
    "Botany": true,
    "Cybersecurity": true,
  };

  return Constants;
})();
